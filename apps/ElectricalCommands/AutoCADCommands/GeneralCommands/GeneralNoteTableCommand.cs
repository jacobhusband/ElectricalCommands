using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const double GeneralNoteNumberWidth = 0.25;
    private const double GeneralNoteTableBaseMargin = 0.025;
    private const double GeneralNoteTableVerticalMargin = 0.125;
    private const double GeneralNoteTableColumnGap = 0.10;
    private const double GeneralNoteRotationTolerance = Math.PI / 180.0;
    private const double GeneralNoteDefaultAttributeHeight = 0.09375;

    [CommandMethod("GENERALNOTETABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("GNTABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("GNT", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void CreateGeneralNoteTable()
    {
      Document document = Application.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        return;
      }

      Database database = document.Database;
      Editor editor = document.Editor;
      PromptSelectionOptions selectionOptions = new PromptSelectionOptions
      {
        MessageForAdding =
          "\nSelect MText note column(s), DBText lines, or an existing General Note Table: ",
        MessageForRemoval = "\nRemove objects: ",
        RejectObjectsFromNonCurrentSpace = true,
      };
      SelectionFilter selectionFilter = new SelectionFilter(
        new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT,ACAD_TABLE") });
      PromptSelectionResult selectionResult = editor.GetSelection(
        selectionOptions,
        selectionFilter);
      if (selectionResult.Status != PromptStatus.OK ||
          selectionResult.Value == null ||
          selectionResult.Value.Count == 0)
      {
        editor.WriteMessage("\nGENERALNOTETABLE canceled.");
        return;
      }

      ObjectId[] sourceIds = selectionResult.Value
        .GetObjectIds()
        .Distinct()
        .ToArray();

      ObjectId existingTableId = FindSelectedTableId(database, sourceIds);
      if (!existingTableId.IsNull)
      {
        UpdateExistingGeneralNoteTable(database, editor, existingTableId);
        return;
      }

      bool useEveryDbTextLine = false;
      bool containsMText;
      bool containsDbText;
      try
      {
        GetGeneralNoteTableSelectionTypes(
          database,
          sourceIds,
          out containsMText,
          out containsDbText);
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nGENERALNOTETABLE selection error: {ex.Message}");
        return;
      }

      if (containsMText && containsDbText)
      {
        editor.WriteMessage(
          "\nSelect either MText columns or DBText lines in one run; " +
          "mixed selections are not converted together.");
        return;
      }

      if (containsDbText)
      {
        PromptKeywordOptions groupingOptions = new PromptKeywordOptions(
          "\nGroup DBText into notes [Automatic/EveryLine] <Automatic>: ",
          "Automatic EveryLine")
        {
          AllowNone = true,
        };
        PromptResult groupingResult = editor.GetKeywords(groupingOptions);
        if (groupingResult.Status != PromptStatus.OK &&
            groupingResult.Status != PromptStatus.None)
        {
          editor.WriteMessage("\nGENERALNOTETABLE canceled.");
          return;
        }
        useEveryDbTextLine = groupingResult.Status == PromptStatus.OK &&
          string.Equals(
            groupingResult.StringResult,
            "EveryLine",
            StringComparison.OrdinalIgnoreCase);
      }

      GeneralNoteTablePlan plan;
      try
      {
        using (Transaction transaction =
          database.TransactionManager.StartOpenCloseTransaction())
        {
          plan = containsMText
            ? BuildGeneralNotePlanFromMText(database, transaction, sourceIds)
            : BuildGeneralNotePlanFromDbText(
              database,
              transaction,
              sourceIds,
              useEveryDbTextLine);
          transaction.Commit();
        }
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nGENERALNOTETABLE could not read the notes: {ex.Message}");
        return;
      }

      if (plan == null || plan.Columns.Count == 0 || plan.NoteCount == 0)
      {
        editor.WriteMessage("\nNo nonblank general-note text was found.");
        return;
      }

      editor.WriteMessage(
        $"\nDetected {plan.NoteCount} note{(plan.NoteCount == 1 ? string.Empty : "s")} " +
        $"in {plan.Columns.Count} column{(plan.Columns.Count == 1 ? string.Empty : "s")}.");
      if (containsDbText && !useEveryDbTextLine &&
          plan.NoteCount == plan.DbVisualLineCount && plan.DbVisualLineCount > 1)
      {
        editor.WriteMessage(
          "\nEach DBText line was detected as a separate note; " +
          "no continuation-line spacing or indentation pattern was found.");
      }
      else if (containsDbText && !useEveryDbTextLine && plan.NoteCount == 1 &&
               plan.DbVisualLineCount > 1)
      {
        editor.WriteMessage(
          "\nWarning: all DBText lines were grouped as one note. " +
          "The source geometry does not contain a detectable note-break pattern.");
      }

      PromptIntegerOptions numberOptions = new PromptIntegerOptions(
        "\nEnter first general-note number <1>: ")
      {
        AllowNone = true,
        AllowNegative = false,
        AllowZero = false,
        DefaultValue = 1,
        LowerLimit = 1,
        UseDefaultValue = true,
      };
      PromptIntegerResult numberResult = editor.GetInteger(numberOptions);
      if (numberResult.Status != PromptStatus.OK &&
          numberResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nGENERALNOTETABLE canceled.");
        return;
      }
      int firstNumber = numberResult.Status == PromptStatus.OK
        ? numberResult.Value
        : 1;

      PromptKeywordOptions layoutOptions = new PromptKeywordOptions(
        "\nTable placement [Boundary/Source] <Boundary>: ",
        "Boundary Source")
      {
        AllowNone = true,
      };
      PromptResult layoutResult = editor.GetKeywords(layoutOptions);
      if (layoutResult.Status != PromptStatus.OK &&
          layoutResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nGENERALNOTETABLE canceled.");
        return;
      }

      bool useBoundary = layoutResult.Status == PromptStatus.None ||
        string.Equals(
          layoutResult.StringResult,
          "Boundary",
          StringComparison.OrdinalIgnoreCase);
      Point3d insertionPoint;
      if (useBoundary)
      {
        if (!TryPromptGeneralNoteBoundary(
          editor,
          plan,
          out GeneralNoteTablePlan boundaryPlan,
          out insertionPoint))
        {
          editor.WriteMessage("\nGENERALNOTETABLE canceled.");
          return;
        }
        plan = boundaryPlan;
      }
      else
      {
        Point3d sourcePosition = GetGeneralNoteTableDefaultPosition(plan);
        PromptPointOptions pointOptions = new PromptPointOptions(
          "\nSpecify table upper-left corner <replace source in place>: ")
        {
          AllowNone = true,
        };
        PromptPointResult pointResult = editor.GetPoint(pointOptions);
        if (pointResult.Status != PromptStatus.OK &&
            pointResult.Status != PromptStatus.None)
        {
          editor.WriteMessage("\nGENERALNOTETABLE canceled.");
          return;
        }
        insertionPoint = pointResult.Status == PromptStatus.OK
          ? pointResult.Value
          : sourcePosition;
      }

      PromptKeywordOptions sourceOptions = new PromptKeywordOptions(
        "\nAfter creating the table [Keep/Erase] source text <Keep>: ",
        "Keep Erase")
      {
        AllowNone = true,
      };
      PromptResult sourceResult = editor.GetKeywords(sourceOptions);
      if (sourceResult.Status != PromptStatus.OK &&
          sourceResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nGENERALNOTETABLE canceled.");
        return;
      }
      bool eraseSource = sourceResult.Status == PromptStatus.OK &&
        string.Equals(
          sourceResult.StringResult,
          "Erase",
          StringComparison.OrdinalIgnoreCase);

      try
      {
        ObjectId tableId = WriteGeneralNoteTable(
          database,
          plan,
          insertionPoint,
          firstNumber,
          sourceIds,
          eraseSource);
        editor.WriteMessage(
          $"\nCreated general-note table {tableId.Handle} with {plan.NoteCount} " +
          $"sequential note{(plan.NoteCount == 1 ? string.Empty : "s")}." +
          (eraseSource ? " Source text erased." : " Source text retained."));
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nGENERALNOTETABLE error: {ex.Message}");
      }
    }

    private static void UpdateExistingGeneralNoteTable(
      Database database,
      Editor editor,
      ObjectId tableId)
    {
      try
      {
        using (Transaction transaction = database.TransactionManager.StartTransaction())
        {
          Table table = transaction.GetObject(tableId, OpenMode.ForWrite) as Table;
          if (table == null)
          {
            editor.WriteMessage("\nSelected object is not a table.");
            return;
          }

          double textHeight = database.Textsize > 1e-9 ? database.Textsize : GeneralNoteDefaultAttributeHeight;
          ObjectId textStyleId = database.Textstyle;
          Color contentColor = Color.FromColorIndex(ColorMethod.ByAci, 256);
          bool hasDot = true;

          // Scan column 0 for existing numbers and text styling
          var foundNumbers = new List<(int row, int number)>();
          for (int row = 0; row < table.Rows.Count; row++)
          {
            if (TryGetCellGeneralNoteNumber(
              table,
              transaction,
              row,
              0,
              out int number,
              out bool cellHasDot,
              out double cellHeight,
              out ObjectId cellStyleId,
              out Color cellColor))
            {
              foundNumbers.Add((row, number));
              hasDot = cellHasDot;
              if (cellHeight > 1e-9)
              {
                textHeight = cellHeight;
              }
              if (!cellStyleId.IsNull && cellStyleId.IsValid)
              {
                textStyleId = cellStyleId;
              }
              if (cellColor != null)
              {
                contentColor = cellColor;
              }
            }
          }

          // Determine numbering step and first row
          int step = 1;
          int lastFoundRow = -1;
          int lastFoundNum = 0;
          int firstDataRow = 0;

          if (foundNumbers.Count >= 2)
          {
            var prev = foundNumbers[foundNumbers.Count - 2];
            var last = foundNumbers[foundNumbers.Count - 1];
            int rowDiff = last.row - prev.row;
            int numDiff = last.number - prev.number;
            if (rowDiff > 0 && numDiff > 0 && numDiff % rowDiff == 0)
            {
              step = numDiff / rowDiff;
            }
            else if (rowDiff > 0 && numDiff > 0)
            {
              step = Math.Max(1, numDiff / rowDiff);
            }
            lastFoundRow = last.row;
            lastFoundNum = last.number;
            firstDataRow = foundNumbers[0].row;
          }
          else if (foundNumbers.Count == 1)
          {
            lastFoundRow = foundNumbers[0].row;
            lastFoundNum = foundNumbers[0].number;
            firstDataRow = foundNumbers[0].row;
          }
          else
          {
            for (int r = 0; r < table.Rows.Count; r++)
            {
              if (!table.IsMergedCell(r, 0, out _))
              {
                firstDataRow = r;
                break;
              }
            }
            lastFoundRow = firstDataRow - 1;
            lastFoundNum = 0;
          }

          int addedCount = 0;
          int firstAdded = 0;
          int lastAdded = 0;

          for (int row = firstDataRow; row < table.Rows.Count; row++)
          {
            bool alreadyHasNumber = foundNumbers.Any(f => f.row == row);
            if (!alreadyHasNumber)
            {
              if (table.IsMergedCell(row, 0, out CellRange range) && range.RightColumn > range.LeftColumn)
              {
                continue;
              }

              int currentNum;
              if (row > lastFoundRow && lastFoundRow >= 0)
              {
                currentNum = lastFoundNum + (row - lastFoundRow) * step;
              }
              else if (foundNumbers.Count > 0)
              {
                currentNum = foundNumbers[0].number + (row - foundNumbers[0].row) * step;
              }
              else
              {
                currentNum = 1 + (row - firstDataRow) * step;
              }

              string numberString = hasDot
                ? $"{currentNum.ToString(CultureInfo.InvariantCulture)}."
                : currentNum.ToString(CultureInfo.InvariantCulture);

              EnsureGeneralNoteTableCellContent(table, row, 0);
              table.SetTextString(row, 0, 0, numberString);
              table.SetTextHeight(row, 0, 0, textHeight);
              table.SetTextStyleId(row, 0, 0, textStyleId);
              table.Cells[row, 0].Alignment = CellAlignment.TopLeft;
              table.SetContentColor(row, 0, 0, contentColor);
              HideGeneralNoteCellGrid(table, row, 0);
              SetGeneralNoteCellMargins(
                table,
                row,
                0,
                GeneralNoteTableVerticalMargin,
                GeneralNoteTableVerticalMargin,
                GeneralNoteTableBaseMargin);

              if (table.Columns.Count > 1)
              {
                HideGeneralNoteCellGrid(table, row, 1);
                EnsureGeneralNoteTableCellContent(table, row, 1);
                SetGeneralNoteCellMargins(
                  table,
                  row,
                  1,
                  GeneralNoteTableVerticalMargin,
                  GeneralNoteTableVerticalMargin,
                  GeneralNoteTableBaseMargin);
              }

              double minRowHeight = textHeight * 1.5 + GeneralNoteTableVerticalMargin * 2.0;
              try
              {
                if (table.Rows[row].Height < minRowHeight)
                {
                  table.Rows[row].Height = minRowHeight;
                }
              }
              catch
              {
              }

              if (addedCount == 0)
              {
                firstAdded = currentNum;
              }
              lastAdded = currentNum;
              addedCount++;
            }
          }

          table.VerticalCellMargin = GeneralNoteTableVerticalMargin;
          table.GenerateLayout();
          TryRecomputeTableBlock(table);

          transaction.Commit();

          if (addedCount > 0)
          {
            editor.WriteMessage(
              $"\nUpdated general-note table {tableId.Handle}: added note number{(addedCount == 1 ? string.Empty : "s")} " +
              $"{firstAdded}{(addedCount > 1 ? $" to {lastAdded}" : string.Empty)} to column 1.");
          }
          else
          {
            string noteRange = foundNumbers.Count > 0
              ? $" (notes {foundNumbers.First().number}" + (foundNumbers.Count > 1 ? $" through {foundNumbers.Last().number}" : string.Empty) + ")"
              : string.Empty;
            editor.WriteMessage($"\nGeneral-note table {tableId.Handle} is already up to date{noteRange}.");
          }
        }
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nGENERALNOTETABLE update error: {ex.Message}");
      }
    }

    private static bool TryGetCellGeneralNoteNumber(
      Table table,
      Transaction transaction,
      int row,
      int column,
      out int number,
      out bool hasDot,
      out double textHeight,
      out ObjectId textStyleId,
      out Color contentColor)
    {
      number = 0;
      hasDot = true;
      textHeight = 0.09375;
      textStyleId = ObjectId.Null;
      contentColor = Color.FromColorIndex(ColorMethod.ByAci, 256);

      try
      {
        if (table.IsMergedCell(row, column, out CellRange range) && range.RightColumn > range.LeftColumn)
        {
          return false;
        }

        string rawText = string.Empty;
        int numContents = table.GetNumberOfContents(row, column);
        for (int i = 0; i < numContents; i++)
        {
          string txt = table.GetTextString(row, column, i);
          if (!string.IsNullOrWhiteSpace(txt))
          {
            rawText = txt;
            try
            {
              textHeight = table.GetTextHeight(row, column, i);
              textStyleId = table.GetTextStyleId(row, column, i);
              contentColor = table.GetContentColor(row, column, i);
            }
            catch
            {
            }
            break;
          }
        }

        if (string.IsNullOrWhiteSpace(rawText))
        {
          rawText = table.Cells[row, column].TextString ?? string.Empty;
          if (textHeight <= 1e-9)
          {
            try
            {
              if (table.Cells[row, column].TextHeight.HasValue)
              {
                textHeight = table.Cells[row, column].TextHeight.Value;
              }
              if (table.Cells[row, column].TextStyleId.HasValue)
              {
                textStyleId = table.Cells[row, column].TextStyleId.Value;
              }
              contentColor = table.Cells[row, column].ContentColor ?? contentColor;
            }
            catch
            {
            }
          }
        }

        if (string.IsNullOrWhiteSpace(rawText))
        {
          return false;
        }

        string plain = RemoveGeneralMTextFormatting(rawText).Trim();
        Match match = Regex.Match(plain, @"^\s*(\d+)");
        if (match.Success && int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
          number = parsed;
          hasDot = plain.Contains(".");
          return true;
        }
      }
      catch
      {
      }

      return false;
    }

    private static bool TryPromptGeneralNoteBoundary(
      Editor editor,
      GeneralNoteTablePlan sourcePlan,
      out GeneralNoteTablePlan boundaryPlan,
      out Point3d insertionPoint)
    {
      boundaryPlan = null;
      insertionPoint = Point3d.Origin;

      PromptPointResult topLeftResult = editor.GetPoint(
        "\nClick the upper-left corner of the first general-note column: ");
      if (topLeftResult.Status != PromptStatus.OK)
      {
        return false;
      }

      PromptCornerOptions cornerOptions = new PromptCornerOptions(
        "\nClick the lower-right corner to set column width and maximum height: ",
        topLeftResult.Value);
      PromptPointResult bottomRightResult = editor.GetCorner(cornerOptions);
      if (bottomRightResult.Status != PromptStatus.OK)
      {
        return false;
      }

      PromptPointOptions directionOptions = new PromptPointOptions(
        "\nClick left or right of the lower-right corner for overflow direction: ")
      {
        BasePoint = bottomRightResult.Value,
        UseBasePoint = true,
        UseDashedLine = true,
      };
      PromptPointResult directionResult = editor.GetPoint(directionOptions);
      if (directionResult.Status != PromptStatus.OK)
      {
        return false;
      }

      GetGeneralNoteLocalCoordinates(
        topLeftResult.Value,
        sourcePlan.Rotation,
        out double topLeftX,
        out double topLeftY);
      GetGeneralNoteLocalCoordinates(
        bottomRightResult.Value,
        sourcePlan.Rotation,
        out double bottomRightX,
        out double bottomRightY);
      GetGeneralNoteLocalCoordinates(
        directionResult.Value,
        sourcePlan.Rotation,
        out double directionX,
        out _);

      double width = bottomRightX - topLeftX;
      double height = topLeftY - bottomRightY;
      if (width <= 1e-6 || height <= 1e-6)
      {
        editor.WriteMessage(
          "\nThe second point must be below and to the right of the first point.");
        return false;
      }

      var singleColumn = new GeneralNoteSourceColumn
      {
        SourceLeft = topLeftX,
        SourceTop = topLeftY,
        SourceWidth = width,
        DbVisualLineCount = sourcePlan.DbVisualLineCount,
      };
      foreach (GeneralNoteSourceColumn sourceColumn in sourcePlan.Columns)
      {
        singleColumn.Entries.AddRange(sourceColumn.Entries);
      }

      double numberWidth = GetGeneralNoteNumberColumnWidth(singleColumn);
      double minimumTextWidth = singleColumn.Entries.Max(
        entry => entry.TextHeight * 2.0);
      if (width <= numberWidth + minimumTextWidth)
      {
        editor.WriteMessage(
          $"\nThe selected column is too narrow. Specify a width greater than " +
          $"{(numberWidth + minimumTextWidth).ToString("0.###", CultureInfo.InvariantCulture)}.");
        return false;
      }

      double scale = GetGeneralNoteColumnScale(singleColumn);
      double maxTextHeight = singleColumn.Entries.Max(entry => entry.TextHeight);
      double minimumRowHeight =
        maxTextHeight +
        (GeneralNoteTableVerticalMargin * 2.0) * scale;
      if (height <= minimumRowHeight)
      {
        editor.WriteMessage(
          $"\nThe selected height is too small to contain one general-note row. " +
          $"Specify a height greater than " +
          $"{minimumRowHeight.ToString("0.###", CultureInfo.InvariantCulture)}.");
        return false;
      }

      TableBreakFlowDirection overflowDirection = directionX < bottomRightX
        ? TableBreakFlowDirection.Left
        : TableBreakFlowDirection.Right;
      boundaryPlan = new GeneralNoteTablePlan
      {
        Columns = new List<GeneralNoteSourceColumn> { singleColumn },
        Rotation = sourcePlan.Rotation,
        Elevation = topLeftResult.Value.Z,
        Layer = sourcePlan.Layer,
        DbVisualLineCount = sourcePlan.DbVisualLineCount,
        ForcedPairWidth = width,
        BreakEnabled = true,
        BreakHeight = height,
        BreakFlowDirection = overflowDirection,
        BreakSpacing = GeneralNoteTableColumnGap * scale,
      };
      insertionPoint = topLeftResult.Value;
      editor.WriteMessage(
        $"\nOverflow will continue to the " +
        $"{(overflowDirection == TableBreakFlowDirection.Left ? "left" : "right")}.");
      return true;
    }

    private static void GetGeneralNoteLocalCoordinates(
      Point3d point,
      double rotation,
      out double x,
      out double y)
    {
      double cosine = Math.Cos(rotation);
      double sine = Math.Sin(rotation);
      x = point.X * cosine + point.Y * sine;
      y = -point.X * sine + point.Y * cosine;
    }

    private static void GetGeneralNoteTableSelectionTypes(
      Database database,
      IEnumerable<ObjectId> sourceIds,
      out bool containsMText,
      out bool containsDbText)
    {
      containsMText = false;
      containsDbText = false;
      using (Transaction transaction =
        database.TransactionManager.StartOpenCloseTransaction())
      {
        foreach (ObjectId sourceId in sourceIds)
        {
          Entity entity = transaction.GetObject(
            sourceId,
            OpenMode.ForRead,
            false) as Entity;
          containsMText |= entity is MText;
          containsDbText |= entity is DBText;
        }
        transaction.Commit();
      }
    }

    private static GeneralNoteTablePlan BuildGeneralNotePlanFromMText(
      Database database,
      Transaction transaction,
      IEnumerable<ObjectId> sourceIds)
    {
      List<MText> sourceTexts = sourceIds
        .Select(id => transaction.GetObject(id, OpenMode.ForRead, false) as MText)
        .Where(text => text != null && !string.IsNullOrWhiteSpace(text.Text))
        .ToList();
      if (sourceTexts.Count == 0)
      {
        return new GeneralNoteTablePlan();
      }

      double rotation = sourceTexts[0].Rotation;
      ValidateGeneralNoteRotations(
        sourceTexts.Select(text => text.Rotation),
        rotation);
      var columns = new List<GeneralNoteSourceColumn>();
      foreach (MText source in sourceTexts)
      {
        GeneralLocalBounds bounds = GetGeneralNoteLocalBounds(source, rotation);
        List<string> paragraphs = SplitMTextIntoGeneralNotes(source);
        if (paragraphs.Count == 0)
        {
          continue;
        }

        double textHeight = PositiveOrDefaultGeneralNote(source.TextHeight, database.Textsize);
        double sourceWidth = source.Width > textHeight
          ? source.Width
          : PositiveOrDefaultGeneralNote(source.ActualWidth, bounds.Width);
        var column = new GeneralNoteSourceColumn
        {
          SourceLeft = bounds.Left,
          SourceTop = bounds.Top,
          SourceWidth = Math.Max(sourceWidth, textHeight * 4.0),
        };
        foreach (string paragraph in paragraphs)
        {
          string cleanedText = StripLeadingNoteNumber(paragraph);
          if (string.IsNullOrWhiteSpace(cleanedText))
          {
            continue;
          }

          column.Entries.Add(new GeneralNoteEntry
          {
            PlainText = cleanedText,
            TextHeight = textHeight,
            TextStyleId = source.TextStyleId,
            ColorIndex = (short)source.ColorIndex,
          });
        }
        columns.Add(column);
      }

      columns = columns.OrderBy(column => column.SourceLeft).ToList();
      return new GeneralNoteTablePlan
      {
        Columns = columns,
        Rotation = rotation,
        Elevation = sourceTexts[0].Location.Z,
        Layer = sourceTexts[0].Layer,
      };
    }

    private static GeneralNoteTablePlan BuildGeneralNotePlanFromDbText(
      Database database,
      Transaction transaction,
      IEnumerable<ObjectId> sourceIds,
      bool everyLine)
    {
      List<DBText> sourceTexts = sourceIds
        .Select(id => transaction.GetObject(id, OpenMode.ForRead, false) as DBText)
        .Where(text => text != null && !string.IsNullOrWhiteSpace(text.TextString))
        .ToList();
      if (sourceTexts.Count == 0)
      {
        return new GeneralNoteTablePlan();
      }

      double rotation = sourceTexts[0].Rotation;
      ValidateGeneralNoteRotations(
        sourceTexts.Select(text => text.Rotation),
        rotation);
      List<GeneralDbTextLine> lines = sourceTexts.Select(source =>
      {
        GeneralLocalBounds bounds = GetGeneralNoteLocalBounds(source, rotation);
        return new GeneralDbTextLine
        {
          Text = source.TextString.Trim(),
          Left = bounds.Left,
          Right = bounds.Right,
          Top = bounds.Top,
          Bottom = bounds.Bottom,
          CenterY = (bounds.Top + bounds.Bottom) / 2.0,
          TextHeight = PositiveOrDefaultGeneralNote(source.Height, database.Textsize),
          TextStyleId = source.TextStyleId,
          ColorIndex = (short)source.ColorIndex,
        };
      }).ToList();

      double medianHeight = MedianGeneralNote(lines.Select(line => line.TextHeight));
      double columnTolerance = Math.Max(medianHeight * 8.0, 1e-4);
      var clusters = new List<List<GeneralDbTextLine>>();
      foreach (GeneralDbTextLine line in lines.OrderBy(line => line.Left))
      {
        List<GeneralDbTextLine> nearest = clusters
          .Where(cluster =>
            Math.Abs(MedianGeneralNote(cluster.Select(item => item.Left)) - line.Left) <=
            columnTolerance)
          .OrderBy(cluster =>
            Math.Abs(MedianGeneralNote(cluster.Select(item => item.Left)) - line.Left))
          .FirstOrDefault();
        if (nearest == null)
        {
          nearest = new List<GeneralDbTextLine>();
          clusters.Add(nearest);
        }
        nearest.Add(line);
      }

      List<GeneralNoteSourceColumn> columns = clusters
        .Select(cluster => BuildGeneralNoteColumnFromDbText(cluster, everyLine))
        .Where(column => column.Entries.Count > 0)
        .OrderBy(column => column.SourceLeft)
        .ToList();
      return new GeneralNoteTablePlan
      {
        Columns = columns,
        Rotation = rotation,
        Elevation = sourceTexts[0].Position.Z,
        Layer = sourceTexts[0].Layer,
        DbVisualLineCount = columns.Sum(column => column.DbVisualLineCount),
      };
    }

    private static GeneralNoteSourceColumn BuildGeneralNoteColumnFromDbText(
      List<GeneralDbTextLine> sourceLines,
      bool everyLine)
    {
      double medianHeight = MedianGeneralNote(sourceLines.Select(line => line.TextHeight));
      double sameLineTolerance = Math.Max(medianHeight * 0.45, 1e-5);
      var visualLines = new List<GeneralDbVisualLine>();
      foreach (GeneralDbTextLine source in sourceLines.OrderByDescending(line => line.CenterY))
      {
        GeneralDbVisualLine visual = visualLines
          .Where(line => Math.Abs(line.CenterY - source.CenterY) <= sameLineTolerance)
          .OrderBy(line => Math.Abs(line.CenterY - source.CenterY))
          .FirstOrDefault();
        if (visual == null)
        {
          visual = new GeneralDbVisualLine();
          visualLines.Add(visual);
        }
        visual.Parts.Add(source);
        visual.Refresh();
      }
      visualLines = visualLines.OrderByDescending(line => line.CenterY).ToList();

      var column = new GeneralNoteSourceColumn
      {
        SourceLeft = visualLines.Min(line => line.Left),
        SourceTop = visualLines.Max(line => line.Top),
        SourceWidth = Math.Max(
          visualLines.Max(line => line.Right) - visualLines.Min(line => line.Left),
          medianHeight * 4.0),
        DbVisualLineCount = visualLines.Count,
      };
      if (everyLine || visualLines.Count == 1)
      {
        foreach (GeneralDbVisualLine line in visualLines)
        {
          column.Entries.Add(CreateGeneralNoteEntry(line, line.Text));
        }
        return column;
      }

      List<double> gaps = new List<double>();
      for (int index = 1; index < visualLines.Count; index++)
      {
        double gap = visualLines[index - 1].CenterY - visualLines[index].CenterY;
        if (gap > 1e-6)
        {
          gaps.Add(gap);
        }
      }
      List<double> lowerGaps = gaps
        .OrderBy(value => value)
        .Take(Math.Max(1, (int)Math.Ceiling(gaps.Count * 0.65)))
        .ToList();
      double normalPitch = lowerGaps.Count > 0
        ? MedianGeneralNote(lowerGaps)
        : medianHeight * 1.5;
      double noteGapThreshold = Math.Max(
        normalPitch * 1.35,
        medianHeight * 1.65);
      double indentationThreshold = Math.Max(medianHeight * 1.25, 1e-5);

      GeneralDbVisualLine first = visualLines[0];
      StringBuilder noteText = new StringBuilder(first.Text);
      GeneralDbVisualLine noteFirstLine = first;
      for (int index = 1; index < visualLines.Count; index++)
      {
        GeneralDbVisualLine previous = visualLines[index - 1];
        GeneralDbVisualLine current = visualLines[index];
        double pitch = previous.CenterY - current.CenterY;
        bool startsNewNote = pitch > noteGapThreshold ||
          current.Left < previous.Left - indentationThreshold;
        if (startsNewNote)
        {
          column.Entries.Add(CreateGeneralNoteEntry(noteFirstLine, noteText.ToString()));
          noteFirstLine = current;
          noteText.Clear();
          noteText.Append(current.Text);
        }
        else
        {
          noteText.Append('\n');
          noteText.Append(current.Text);
        }
      }
      column.Entries.Add(CreateGeneralNoteEntry(noteFirstLine, noteText.ToString()));
      return column;
    }

    private static GeneralNoteEntry CreateGeneralNoteEntry(GeneralDbVisualLine line, string text)
    {
      return new GeneralNoteEntry
      {
        PlainText = StripLeadingNoteNumber(text),
        TextHeight = line.TextHeight,
        TextStyleId = line.TextStyleId,
        ColorIndex = line.ColorIndex,
      };
    }

    private static string StripLeadingNoteNumber(string text)
    {
      if (string.IsNullOrWhiteSpace(text))
      {
        return string.Empty;
      }
      return Regex.Replace(
        text,
        @"^\s*(?:\(?\d+\)?\s*[\.\:\)]|\d+\.)\s*(?:\-\s*)?",
        string.Empty);
    }

    private static List<string> SplitMTextIntoGeneralNotes(MText source)
    {
      string visible = (source.Text ?? string.Empty)
        .Replace("\r\n", "\n")
        .Replace('\r', '\n')
        .Replace("\\P", "\n");
      List<string> visibleParagraphs = visible
        .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(value => value.Trim())
        .Where(value => value.Length > 0)
        .ToList();

      int rawParagraphCount = CountGeneralMTextParagraphBreaks(source.Contents) + 1;
      if (rawParagraphCount <= 1)
      {
        return visibleParagraphs.Count == 0
          ? new List<string>()
          : new List<string> { string.Join("\n", visibleParagraphs) };
      }
      if (visibleParagraphs.Count == rawParagraphCount)
      {
        return visibleParagraphs;
      }

      return SplitRawGeneralMTextParagraphs(source.Contents)
        .Select(RemoveGeneralMTextFormatting)
        .Select(value => value.Trim())
        .Where(value => value.Length > 0)
        .ToList();
    }

    private static int CountGeneralMTextParagraphBreaks(string contents)
    {
      return SplitRawGeneralMTextParagraphs(contents).Count - 1;
    }

    private static List<string> SplitRawGeneralMTextParagraphs(string contents)
    {
      var paragraphs = new List<string>();
      var current = new StringBuilder();
      string value = contents ?? string.Empty;
      for (int index = 0; index < value.Length; index++)
      {
        char ch = value[index];
        if (ch == '\\' && index + 1 < value.Length)
        {
          char next = value[index + 1];
          if (next == '\\')
          {
            current.Append(ch);
            current.Append(next);
            index++;
            continue;
          }
          if (next == 'P')
          {
            paragraphs.Add(current.ToString());
            current.Clear();
            index++;
            continue;
          }
        }
        if (ch == '\r' || ch == '\n')
        {
          paragraphs.Add(current.ToString());
          current.Clear();
          if (ch == '\r' && index + 1 < value.Length && value[index + 1] == '\n')
          {
            index++;
          }
          continue;
        }
        current.Append(ch);
      }
      paragraphs.Add(current.ToString());
      return paragraphs;
    }

    private static string RemoveGeneralMTextFormatting(string contents)
    {
      var result = new StringBuilder();
      string value = contents ?? string.Empty;
      for (int index = 0; index < value.Length; index++)
      {
        char ch = value[index];
        if (ch == '{' || ch == '}')
        {
          continue;
        }
        if (ch != '\\' || index + 1 >= value.Length)
        {
          result.Append(ch);
          continue;
        }

        char code = value[++index];
        if (code == '\\' || code == '{' || code == '}')
        {
          result.Append(code);
        }
        else if (code == '~')
        {
          result.Append(' ');
        }
        else if (code == 'n')
        {
          result.Append('\n');
        }
        else if (code == 'U' && index + 5 < value.Length && value[index + 1] == '+')
        {
          string hex = value.Substring(index + 2, 4);
          if (int.TryParse(
            hex,
            NumberStyles.HexNumber,
            CultureInfo.InvariantCulture,
            out int unicode))
          {
            result.Append((char)unicode);
            index += 5;
          }
        }
        else if (code == 'S')
        {
          int end = value.IndexOf(';', index + 1);
          if (end < 0)
          {
            continue;
          }
          string stacked = value.Substring(index + 1, end - index - 1)
            .Replace('#', '/')
            .Replace('^', '/');
          result.Append(stacked);
          index = end;
        }
        else if ("AaCcFfHhQqTtWwp".IndexOf(code) >= 0)
        {
          int end = value.IndexOf(';', index + 1);
          index = end < 0 ? value.Length : end;
        }
        // Toggle codes (underline, overline, strike-through) carry no text.
        else if ("LlOoKkXx".IndexOf(code) < 0)
        {
          result.Append(code);
        }
      }
      return result.ToString();
    }

    private static ObjectId WriteGeneralNoteTable(
      Database database,
      GeneralNoteTablePlan plan,
      Point3d insertionPoint,
      int firstNumber,
      IEnumerable<ObjectId> sourceIds,
      bool eraseSource)
    {
      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        Table table = BuildGeneralNoteTable(
          database,
          plan,
          insertionPoint,
          firstNumber);
        BlockTableRecord space = (BlockTableRecord)transaction.GetObject(
          database.CurrentSpaceId,
          OpenMode.ForWrite);
        ObjectId tableId = space.AppendEntity(table);
        transaction.AddNewlyCreatedDBObject(table, true);
        table.GenerateLayout();
        EnsureMinimumGeneralNoteRowHeights(table, plan);
        table.GenerateLayout();
        ConfigureGeneralNoteTableBreaks(table, plan);
        table.GenerateLayout();
        TryRecomputeTableBlock(table);

        if (eraseSource)
        {
          foreach (ObjectId sourceId in sourceIds)
          {
            Entity source = transaction.GetObject(
              sourceId,
              OpenMode.ForWrite,
              false) as Entity;
            source?.Erase();
          }
        }
        transaction.Commit();
        return tableId;
      }
    }

    private static void ConfigureGeneralNoteTableBreaks(
      Table table,
      GeneralNoteTablePlan plan)
    {
      if (!plan.BreakEnabled)
      {
        return;
      }

      try
      {
        table.BreakOptions =
          TableBreakOptions.EnableBreaking |
          TableBreakOptions.AllowManualHeights;
        table.BreakEnabled = true;
        table.BreakFlowDirection = plan.BreakFlowDirection;
        table.SetBreakSpacing(Math.Max(plan.BreakSpacing, 0.0));
        table.SetBreakHeight(0, plan.BreakHeight);
      }
      catch (System.Exception ex)
      {
        throw new InvalidOperationException(
          $"Unable to configure table overflow: {ex.Message}",
          ex);
      }
    }

    private static Table BuildGeneralNoteTable(
      Database database,
      GeneralNoteTablePlan plan,
      Point3d insertionPoint,
      int firstNumber)
    {
      int rowCount = plan.Columns.Max(column => column.Entries.Count);
      int columnCount = plan.Columns.Count * 2;
      var table = new Table();
      string stage = "initializing the table";
      try
      {
        stage = "applying database defaults";
        table.SetDatabaseDefaults(database);
        stage = "assigning the current table style";
        table.TableStyle = database.Tablestyle;
        stage = "setting the table position";
        table.Position = insertionPoint;
        stage = "setting the table rotation";
        table.Rotation = plan.Rotation;
        stage = "creating the table rows and columns";
        table.SetSize(rowCount, columnCount);
        stage = "removing table-style title and header merges";
        UnmergeGeneralNoteTableCells(table);
        table.VerticalCellMargin = GeneralNoteTableVerticalMargin;
        if (!string.IsNullOrWhiteSpace(plan.Layer))
        {
          stage = $"assigning table layer '{plan.Layer}'";
          table.Layer = plan.Layer;
        }

        stage = "sizing table columns";
        double[] numberWidths = plan.Columns
          .Select(GetGeneralNoteNumberColumnWidth)
          .ToArray();
        for (int columnIndex = 0; columnIndex < plan.Columns.Count; columnIndex++)
        {
          GeneralNoteSourceColumn column = plan.Columns[columnIndex];
          int numberColumn = columnIndex * 2;
          int textColumn = numberColumn + 1;
          table.Columns[numberColumn].Width = numberWidths[columnIndex];
          double scale = GetGeneralNoteColumnScale(column);
          double margin = GeneralNoteTableBaseMargin * scale;
          double contentWidth;
          if (plan.ForcedPairWidth > 1e-9)
          {
            contentWidth = plan.ForcedPairWidth - numberWidths[columnIndex];
            if (contentWidth <= margin * 2.0)
            {
              throw new InvalidOperationException(
                "The selected table width leaves no usable space for note text.");
            }
          }
          else
          {
            contentWidth = Math.Max(
              column.SourceWidth + margin * 2.0,
              column.Entries.Max(entry => entry.TextHeight) * 4.0);
            if (columnIndex + 1 < plan.Columns.Count)
            {
              double sourceDelta = plan.Columns[columnIndex + 1].SourceLeft -
                column.SourceLeft;
              double alignmentWidth = sourceDelta - numberWidths[columnIndex + 1];
              contentWidth = Math.Max(contentWidth, alignmentWidth);
            }
            else
            {
              contentWidth += GeneralNoteTableColumnGap * scale;
            }
          }
          table.Columns[textColumn].Width = contentWidth;
        }

        int number = firstNumber;
        for (int sourceColumn = 0; sourceColumn < plan.Columns.Count; sourceColumn++)
        {
          GeneralNoteSourceColumn column = plan.Columns[sourceColumn];
          int numberColumn = sourceColumn * 2;
          int textColumn = numberColumn + 1;
          for (int row = 0; row < rowCount; row++)
          {
            stage = $"formatting row {row + 1}, column {sourceColumn + 1}";
            HideGeneralNoteCellGrid(table, row, numberColumn);
            HideGeneralNoteCellGrid(table, row, textColumn);
            EnsureGeneralNoteTableCellContent(table, row, numberColumn);
            EnsureGeneralNoteTableCellContent(table, row, textColumn);
            SetGeneralNoteCellMargins(
              table,
              row,
              numberColumn,
              GeneralNoteTableVerticalMargin,
              GeneralNoteTableVerticalMargin,
              GeneralNoteTableBaseMargin);
            SetGeneralNoteCellMargins(
              table,
              row,
              textColumn,
              GeneralNoteTableVerticalMargin,
              GeneralNoteTableVerticalMargin,
              GeneralNoteTableBaseMargin);
            if (row >= column.Entries.Count)
            {
              table.SetTextString(row, numberColumn, 0, string.Empty);
              table.SetTextString(row, textColumn, 0, string.Empty);
              continue;
            }

            GeneralNoteEntry entry = column.Entries[row];

            stage = $"writing general-note number {number}";
            table.SetTextString(
              row,
              numberColumn,
              0,
              $"{number.ToString(CultureInfo.InvariantCulture)}.");
            table.SetTextHeight(row, numberColumn, 0, entry.TextHeight);
            table.SetTextStyleId(row, numberColumn, 0, entry.TextStyleId);
            table.Cells[row, numberColumn].Alignment = CellAlignment.TopLeft;
            table.SetContentColor(
              row,
              numberColumn,
              0,
              ResolveGeneralNoteTableColor(entry.ColorIndex));

            stage = $"writing note text for note {number}";
            table.SetTextString(
              row,
              textColumn,
              0,
              ConvertPlainTextToMTextContents(StripLeadingNoteNumber(entry.PlainText)));
            table.SetTextHeight(row, textColumn, 0, entry.TextHeight);
            table.SetTextStyleId(row, textColumn, 0, entry.TextStyleId);
            table.Cells[row, textColumn].Alignment = CellAlignment.TopLeft;
            table.SetContentColor(
              row,
              textColumn,
              0,
              ResolveGeneralNoteTableColor(entry.ColorIndex));
            number++;
          }
        }

        stage = "generating the initial table layout";
        table.GenerateLayout();
        stage = "sizing table rows";
        EnsureMinimumGeneralNoteRowHeights(table, plan);
        stage = "generating the final table layout";
        table.GenerateLayout();
        TryRecomputeTableBlock(table);
        return table;
      }
      catch (System.Exception ex)
      {
        table.Dispose();
        throw new InvalidOperationException(
          $"Table construction failed while {stage}: {ex.Message}",
          ex);
      }
    }

    private static void EnsureMinimumGeneralNoteRowHeights(
      Table table,
      GeneralNoteTablePlan plan)
    {
      if (table == null || plan == null)
      {
        return;
      }

      for (int row = 0; row < table.Rows.Count; row++)
      {
        try
        {
          double maxTextHeight = GeneralNoteDefaultAttributeHeight;
          foreach (var col in plan.Columns)
          {
            if (row < col.Entries.Count)
            {
              maxTextHeight = Math.Max(maxTextHeight, col.Entries[row].TextHeight);
            }
          }
          double minRowHeight = maxTextHeight * 1.5 + GeneralNoteTableVerticalMargin * 2.0;
          if (table.Rows[row].Height < minRowHeight)
          {
            table.Rows[row].Height = minRowHeight;
          }
        }
        catch
        {
        }
      }
    }

    private static void UnmergeGeneralNoteTableCells(Table table)
    {
      for (int row = 0; row < table.Rows.Count; row++)
      {
        for (int column = 0; column < table.Columns.Count; column++)
        {
          CellRange range;
          if (table.IsMergedCell(row, column, out range))
          {
            table.UnmergeCells(range);
          }
        }
      }
    }

    private static void EnsureGeneralNoteTableCellContent(
      Table table,
      int row,
      int column)
    {
      int count = table.GetNumberOfContents(row, column);
      if (count <= 0)
      {
        table.CreateContent(row, column, 0);
        count = table.GetNumberOfContents(row, column);
      }
      if (count <= 0)
      {
        throw new InvalidOperationException(
          $"AutoCAD did not create table-cell content at row {row + 1}, " +
          $"column {column + 1}.");
      }
    }

    private static Color ResolveGeneralNoteTableColor(short colorIndex)
    {
      if (colorIndex == 0)
      {
        return Color.FromColorIndex(ColorMethod.ByBlock, 0);
      }
      if (colorIndex < 0 || colorIndex >= 256)
      {
        return Color.FromColorIndex(ColorMethod.ByLayer, 256);
      }
      return Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
    }

    private static void HideGeneralNoteCellGrid(Table table, int row, int column)
    {
      try
      {
        CellBorders borders = table.Cells[row, column].Borders;
        borders.Top.IsVisible = false;
        borders.Bottom.IsVisible = false;
        borders.Left.IsVisible = false;
        borders.Right.IsVisible = false;
        borders.Horizontal.IsVisible = false;
        borders.Vertical.IsVisible = false;
        return;
      }
      catch (System.Exception)
      {
        // Fall back to the legacy table API for older AutoCAD releases.
      }

      GridLineType[] gridLines =
      {
        GridLineType.HorizontalTop,
        GridLineType.HorizontalInside,
        GridLineType.HorizontalBottom,
        GridLineType.VerticalLeft,
        GridLineType.VerticalInside,
        GridLineType.VerticalRight,
      };
      foreach (GridLineType gridLine in gridLines)
      {
        try
        {
          table.SetGridVisibility(
            row,
            column,
            gridLine,
            Visibility.Invisible);
        }
        catch (Autodesk.AutoCAD.Runtime.Exception)
        {
          // A boundary-specific grid type can be unavailable for an edge cell.
        }
      }
    }

    private static void SetGeneralNoteCellMargins(
      Table table,
      int row,
      int column,
      double topMargin,
      double bottomMargin,
      double horizontalMargin)
    {
      try
      {
        CellBorders borders = table.Cells[row, column].Borders;
        borders.Top.Margin = topMargin;
        borders.Bottom.Margin = bottomMargin;
        borders.Left.Margin = horizontalMargin;
        borders.Right.Margin = horizontalMargin;
        return;
      }
      catch (System.Exception)
      {
        // Fall back to the legacy table API for older AutoCAD releases.
      }

      table.SetMargin(row, column, CellMargins.Top, topMargin);
      table.SetMargin(row, column, CellMargins.Bottom, bottomMargin);
      table.SetMargin(row, column, CellMargins.Left, horizontalMargin);
      table.SetMargin(row, column, CellMargins.Right, horizontalMargin);
    }

    private static Point3d GetGeneralNoteTableDefaultPosition(GeneralNoteTablePlan plan)
    {
      GeneralNoteSourceColumn first = plan.Columns[0];
      double left = first.SourceLeft - GetGeneralNoteNumberColumnWidth(first);
      double top = plan.Columns.Max(column => column.SourceTop);
      double cosine = Math.Cos(plan.Rotation);
      double sine = Math.Sin(plan.Rotation);
      return new Point3d(
        left * cosine - top * sine,
        left * sine + top * cosine,
        plan.Elevation);
    }

    private static double GetGeneralNoteNumberColumnWidth(GeneralNoteSourceColumn column)
    {
      double scale = GetGeneralNoteColumnScale(column);
      return (GeneralNoteNumberWidth + GeneralNoteTableBaseMargin * 2.0) * scale;
    }

    private static double GetGeneralNoteColumnScale(GeneralNoteSourceColumn column)
    {
      return column.Entries.Count == 0
        ? 1.0
        : column.Entries.Max(GetGeneralNoteEntryScale);
    }

    private static double GetGeneralNoteEntryScale(GeneralNoteEntry entry)
    {
      return Math.Max(entry.TextHeight / GeneralNoteDefaultAttributeHeight, 1e-6);
    }

    private static void ValidateGeneralNoteRotations(
      IEnumerable<double> rotations,
      double reference)
    {
      if (rotations.Any(rotation =>
        AngularDifferenceGeneralNote(rotation, reference) > GeneralNoteRotationTolerance))
      {
        throw new InvalidOperationException(
          "Selected text objects must share the same rotation (within one degree).");
      }
    }

    private static double AngularDifferenceGeneralNote(double first, double second)
    {
      double difference = Math.Abs(first - second) % (Math.PI * 2.0);
      return difference > Math.PI
        ? Math.PI * 2.0 - difference
        : difference;
    }

    private static GeneralLocalBounds GetGeneralNoteLocalBounds(Entity entity, double rotation)
    {
      Extents3d extents = entity.GeometricExtents;
      Point3d[] corners =
      {
        new Point3d(extents.MinPoint.X, extents.MinPoint.Y, extents.MinPoint.Z),
        new Point3d(extents.MinPoint.X, extents.MaxPoint.Y, extents.MinPoint.Z),
        new Point3d(extents.MaxPoint.X, extents.MinPoint.Y, extents.MaxPoint.Z),
        new Point3d(extents.MaxPoint.X, extents.MaxPoint.Y, extents.MaxPoint.Z),
      };
      double cosine = Math.Cos(rotation);
      double sine = Math.Sin(rotation);
      double[] localX = corners
        .Select(point => point.X * cosine + point.Y * sine)
        .ToArray();
      double[] localY = corners
        .Select(point => -point.X * sine + point.Y * cosine)
        .ToArray();
      return new GeneralLocalBounds
      {
        Left = localX.Min(),
        Right = localX.Max(),
        Bottom = localY.Min(),
        Top = localY.Max(),
      };
    }

    private static double PositiveOrDefaultGeneralNote(double value, double fallback)
    {
      if (value > 1e-9)
      {
        return value;
      }
      return fallback > 1e-9 ? fallback : GeneralNoteDefaultAttributeHeight;
    }

    private static double MedianGeneralNote(IEnumerable<double> values)
    {
      double[] ordered = values
        .Where(value => !double.IsNaN(value) && !double.IsInfinity(value))
        .OrderBy(value => value)
        .ToArray();
      if (ordered.Length == 0)
      {
        return 0.0;
      }
      int middle = ordered.Length / 2;
      return ordered.Length % 2 == 0
        ? (ordered[middle - 1] + ordered[middle]) / 2.0
        : ordered[middle];
    }

    private sealed class GeneralNoteTablePlan
    {
      internal List<GeneralNoteSourceColumn> Columns { get; set; } =
        new List<GeneralNoteSourceColumn>();
      internal double Rotation { get; set; }
      internal double Elevation { get; set; }
      internal string Layer { get; set; } = string.Empty;
      internal int DbVisualLineCount { get; set; }
      internal double ForcedPairWidth { get; set; }
      internal bool BreakEnabled { get; set; }
      internal double BreakHeight { get; set; }
      internal double BreakSpacing { get; set; }
      internal TableBreakFlowDirection BreakFlowDirection { get; set; } =
        TableBreakFlowDirection.Right;
      internal int NoteCount => Columns.Sum(column => column.Entries.Count);
    }

    private sealed class GeneralNoteSourceColumn
    {
      internal List<GeneralNoteEntry> Entries { get; } = new List<GeneralNoteEntry>();
      internal double SourceLeft { get; set; }
      internal double SourceTop { get; set; }
      internal double SourceWidth { get; set; }
      internal int DbVisualLineCount { get; set; }
    }

    private sealed class GeneralNoteEntry
    {
      internal string PlainText { get; set; } = string.Empty;
      internal double TextHeight { get; set; }
      internal ObjectId TextStyleId { get; set; }
      internal short ColorIndex { get; set; } = 256;
    }

    private sealed class GeneralDbTextLine
    {
      internal string Text { get; set; } = string.Empty;
      internal double Left { get; set; }
      internal double Right { get; set; }
      internal double Top { get; set; }
      internal double Bottom { get; set; }
      internal double CenterY { get; set; }
      internal double TextHeight { get; set; }
      internal ObjectId TextStyleId { get; set; }
      internal short ColorIndex { get; set; } = 256;
    }

    private sealed class GeneralDbVisualLine
    {
      internal List<GeneralDbTextLine> Parts { get; } = new List<GeneralDbTextLine>();
      internal string Text { get; private set; } = string.Empty;
      internal double Left { get; private set; }
      internal double Right { get; private set; }
      internal double Top { get; private set; }
      internal double CenterY { get; private set; }
      internal double TextHeight { get; private set; }
      internal ObjectId TextStyleId { get; private set; }
      internal short ColorIndex { get; private set; } = 256;

      internal void Refresh()
      {
        List<GeneralDbTextLine> ordered = Parts.OrderBy(part => part.Left).ToList();
        Text = string.Join(" ", ordered.Select(part => part.Text));
        Left = ordered.Min(part => part.Left);
        Right = ordered.Max(part => part.Right);
        Top = ordered.Max(part => part.Top);
        CenterY = MedianGeneralNote(ordered.Select(part => part.CenterY));
        TextHeight = ordered.Max(part => part.TextHeight);
        TextStyleId = ordered[0].TextStyleId;
        ColorIndex = ordered[0].ColorIndex;
      }
    }

    private sealed class GeneralLocalBounds
    {
      internal double Left { get; set; }
      internal double Right { get; set; }
      internal double Bottom { get; set; }
      internal double Top { get; set; }
      internal double Width => Right - Left;
    }
  }
}
