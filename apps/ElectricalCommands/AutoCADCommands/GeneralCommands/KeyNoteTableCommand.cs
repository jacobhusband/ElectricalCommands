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
using System.Text;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const double KeyNoteHexWidth = 0.2504;
    private const double KeyNoteHexHeight = 0.216852;
    private const double KeyNoteTableBaseMargin = 0.025;
    private const double KeyNoteTableColumnGap = 0.10;
    private const double KeyNoteRotationTolerance = Math.PI / 180.0;

    [CommandMethod("KEYNOTETABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("KNTABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("KNT", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void CreateKeyNoteTable()
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
          "\nSelect MText note column(s), or the DBText lines that make up the notes: ",
        MessageForRemoval = "\nRemove text objects: ",
        RejectObjectsFromNonCurrentSpace = true,
      };
      SelectionFilter selectionFilter = new SelectionFilter(
        new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT") });
      PromptSelectionResult selectionResult = editor.GetSelection(
        selectionOptions,
        selectionFilter);
      if (selectionResult.Status != PromptStatus.OK ||
          selectionResult.Value == null ||
          selectionResult.Value.Count == 0)
      {
        editor.WriteMessage("\nKEYNOTETABLE canceled.");
        return;
      }

      ObjectId[] sourceIds = selectionResult.Value
        .GetObjectIds()
        .Distinct()
        .ToArray();
      bool useEveryDbTextLine = false;
      bool containsMText;
      bool containsDbText;
      try
      {
        GetKeyNoteTableSelectionTypes(
          database,
          sourceIds,
          out containsMText,
          out containsDbText);
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nKEYNOTETABLE selection error: {ex.Message}");
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
          editor.WriteMessage("\nKEYNOTETABLE canceled.");
          return;
        }
        useEveryDbTextLine = groupingResult.Status == PromptStatus.OK &&
          string.Equals(
            groupingResult.StringResult,
            "EveryLine",
            StringComparison.OrdinalIgnoreCase);
      }

      KeyNoteTablePlan plan;
      try
      {
        using (Transaction transaction =
          database.TransactionManager.StartOpenCloseTransaction())
        {
          plan = containsMText
            ? BuildKeyNotePlanFromMText(database, transaction, sourceIds)
            : BuildKeyNotePlanFromDbText(
              database,
              transaction,
              sourceIds,
              useEveryDbTextLine);
          transaction.Commit();
        }
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nKEYNOTETABLE could not read the notes: {ex.Message}");
        return;
      }

      if (plan == null || plan.Columns.Count == 0 || plan.NoteCount == 0)
      {
        editor.WriteMessage("\nNo nonblank keyed-note text was found.");
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
        "\nEnter first keyed-note number <1>: ")
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
        editor.WriteMessage("\nKEYNOTETABLE canceled.");
        return;
      }
      int firstNumber = numberResult.Status == PromptStatus.OK
        ? numberResult.Value
        : 1;

      Point3d sourcePosition = GetKeyNoteTableDefaultPosition(plan);
      PromptPointOptions pointOptions = new PromptPointOptions(
        "\nSpecify table upper-left corner <replace source in place>: ")
      {
        AllowNone = true,
      };
      PromptPointResult pointResult = editor.GetPoint(pointOptions);
      if (pointResult.Status != PromptStatus.OK &&
          pointResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nKEYNOTETABLE canceled.");
        return;
      }
      Point3d insertionPoint = pointResult.Status == PromptStatus.OK
        ? pointResult.Value
        : sourcePosition;

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
        editor.WriteMessage("\nKEYNOTETABLE canceled.");
        return;
      }
      bool eraseSource = sourceResult.Status == PromptStatus.OK &&
        string.Equals(
          sourceResult.StringResult,
          "Erase",
          StringComparison.OrdinalIgnoreCase);

      try
      {
        EnsureKnLayer(database);
        ObjectId textStyleId = EnsureKnTextStyle(database);
        if (!EnsureKnBlockDefinition(database, textStyleId))
        {
          editor.WriteMessage(
            $"\nUnable to prepare the canonical {KnBlockName} block.");
          return;
        }

        ObjectId tableId = WriteKeyNoteTable(
          database,
          plan,
          insertionPoint,
          firstNumber,
          sourceIds,
          eraseSource);
        editor.WriteMessage(
          $"\nCreated keyed-note table {tableId.Handle} with {plan.NoteCount} " +
          $"sequential symbol{(plan.NoteCount == 1 ? string.Empty : "s")}." +
          (eraseSource ? " Source text erased." : " Source text retained."));
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nKEYNOTETABLE error: {ex.Message}");
      }
    }

    private static void GetKeyNoteTableSelectionTypes(
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

    private static KeyNoteTablePlan BuildKeyNotePlanFromMText(
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
        return new KeyNoteTablePlan();
      }

      double rotation = sourceTexts[0].Rotation;
      ValidateKeyNoteRotations(
        sourceTexts.Select(text => text.Rotation),
        rotation);
      var columns = new List<KeyNoteSourceColumn>();
      foreach (MText source in sourceTexts)
      {
        LocalBounds bounds = GetLocalBounds(source, rotation);
        List<string> paragraphs = SplitMTextIntoKeyNotes(source);
        if (paragraphs.Count == 0)
        {
          continue;
        }

        double textHeight = PositiveOrDefault(source.TextHeight, database.Textsize);
        double sourceWidth = source.Width > textHeight
          ? source.Width
          : PositiveOrDefault(source.ActualWidth, bounds.Width);
        var column = new KeyNoteSourceColumn
        {
          SourceLeft = bounds.Left,
          SourceTop = bounds.Top,
          SourceWidth = Math.Max(sourceWidth, textHeight * 4.0),
        };
        foreach (string paragraph in paragraphs)
        {
          column.Entries.Add(new KeyNoteEntry
          {
            PlainText = paragraph,
            TextHeight = textHeight,
            TextStyleId = source.TextStyleId,
            ColorIndex = (short)source.ColorIndex,
          });
        }
        columns.Add(column);
      }

      columns = columns.OrderBy(column => column.SourceLeft).ToList();
      return new KeyNoteTablePlan
      {
        Columns = columns,
        Rotation = rotation,
        Elevation = sourceTexts[0].Location.Z,
        Layer = sourceTexts[0].Layer,
      };
    }

    private static KeyNoteTablePlan BuildKeyNotePlanFromDbText(
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
        return new KeyNoteTablePlan();
      }

      double rotation = sourceTexts[0].Rotation;
      ValidateKeyNoteRotations(
        sourceTexts.Select(text => text.Rotation),
        rotation);
      List<DbTextLine> lines = sourceTexts.Select(source =>
      {
        LocalBounds bounds = GetLocalBounds(source, rotation);
        return new DbTextLine
        {
          Text = source.TextString.Trim(),
          Left = bounds.Left,
          Right = bounds.Right,
          Top = bounds.Top,
          Bottom = bounds.Bottom,
          CenterY = (bounds.Top + bounds.Bottom) / 2.0,
          TextHeight = PositiveOrDefault(source.Height, database.Textsize),
          TextStyleId = source.TextStyleId,
          ColorIndex = (short)source.ColorIndex,
        };
      }).ToList();

      double medianHeight = Median(lines.Select(line => line.TextHeight));
      double columnTolerance = Math.Max(medianHeight * 8.0, 1e-4);
      var clusters = new List<List<DbTextLine>>();
      foreach (DbTextLine line in lines.OrderBy(line => line.Left))
      {
        List<DbTextLine> nearest = clusters
          .Where(cluster =>
            Math.Abs(Median(cluster.Select(item => item.Left)) - line.Left) <=
            columnTolerance)
          .OrderBy(cluster =>
            Math.Abs(Median(cluster.Select(item => item.Left)) - line.Left))
          .FirstOrDefault();
        if (nearest == null)
        {
          nearest = new List<DbTextLine>();
          clusters.Add(nearest);
        }
        nearest.Add(line);
      }

      List<KeyNoteSourceColumn> columns = clusters
        .Select(cluster => BuildKeyNoteColumnFromDbText(cluster, everyLine))
        .Where(column => column.Entries.Count > 0)
        .OrderBy(column => column.SourceLeft)
        .ToList();
      return new KeyNoteTablePlan
      {
        Columns = columns,
        Rotation = rotation,
        Elevation = sourceTexts[0].Position.Z,
        Layer = sourceTexts[0].Layer,
        DbVisualLineCount = columns.Sum(column => column.DbVisualLineCount),
      };
    }

    private static KeyNoteSourceColumn BuildKeyNoteColumnFromDbText(
      List<DbTextLine> sourceLines,
      bool everyLine)
    {
      double medianHeight = Median(sourceLines.Select(line => line.TextHeight));
      double sameLineTolerance = Math.Max(medianHeight * 0.45, 1e-5);
      var visualLines = new List<DbVisualLine>();
      foreach (DbTextLine source in sourceLines.OrderByDescending(line => line.CenterY))
      {
        DbVisualLine visual = visualLines
          .Where(line => Math.Abs(line.CenterY - source.CenterY) <= sameLineTolerance)
          .OrderBy(line => Math.Abs(line.CenterY - source.CenterY))
          .FirstOrDefault();
        if (visual == null)
        {
          visual = new DbVisualLine();
          visualLines.Add(visual);
        }
        visual.Parts.Add(source);
        visual.Refresh();
      }
      visualLines = visualLines.OrderByDescending(line => line.CenterY).ToList();

      var column = new KeyNoteSourceColumn
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
        foreach (DbVisualLine line in visualLines)
        {
          column.Entries.Add(CreateKeyNoteEntry(line, line.Text));
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
        ? Median(lowerGaps)
        : medianHeight * 1.5;
      double noteGapThreshold = Math.Max(
        normalPitch * 1.35,
        medianHeight * 1.65);
      double indentationThreshold = Math.Max(medianHeight * 1.25, 1e-5);

      DbVisualLine first = visualLines[0];
      StringBuilder noteText = new StringBuilder(first.Text);
      DbVisualLine noteFirstLine = first;
      for (int index = 1; index < visualLines.Count; index++)
      {
        DbVisualLine previous = visualLines[index - 1];
        DbVisualLine current = visualLines[index];
        double pitch = previous.CenterY - current.CenterY;
        bool startsNewNote = pitch > noteGapThreshold ||
          current.Left < previous.Left - indentationThreshold;
        if (startsNewNote)
        {
          column.Entries.Add(CreateKeyNoteEntry(noteFirstLine, noteText.ToString()));
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
      column.Entries.Add(CreateKeyNoteEntry(noteFirstLine, noteText.ToString()));
      return column;
    }

    private static KeyNoteEntry CreateKeyNoteEntry(DbVisualLine line, string text)
    {
      return new KeyNoteEntry
      {
        PlainText = text,
        TextHeight = line.TextHeight,
        TextStyleId = line.TextStyleId,
        ColorIndex = line.ColorIndex,
      };
    }

    private static List<string> SplitMTextIntoKeyNotes(MText source)
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

      int rawParagraphCount = CountMTextParagraphBreaks(source.Contents) + 1;
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

      return SplitRawMTextParagraphs(source.Contents)
        .Select(RemoveMTextFormatting)
        .Select(value => value.Trim())
        .Where(value => value.Length > 0)
        .ToList();
    }

    private static int CountMTextParagraphBreaks(string contents)
    {
      return SplitRawMTextParagraphs(contents).Count - 1;
    }

    private static List<string> SplitRawMTextParagraphs(string contents)
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

    private static string RemoveMTextFormatting(string contents)
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

    private static ObjectId WriteKeyNoteTable(
      Database database,
      KeyNoteTablePlan plan,
      Point3d insertionPoint,
      int firstNumber,
      IEnumerable<ObjectId> sourceIds,
      bool eraseSource)
    {
      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        BlockTable blockTable = (BlockTable)transaction.GetObject(
          database.BlockTableId,
          OpenMode.ForRead);
        if (!blockTable.Has(KnBlockName))
        {
          throw new InvalidOperationException($"Block {KnBlockName} is missing.");
        }
        ObjectId blockId = blockTable[KnBlockName];
        ObjectId attributeDefinitionId = FindKeyNoteAttributeDefinition(
          transaction,
          blockId);
        if (attributeDefinitionId.IsNull)
        {
          throw new InvalidOperationException(
            $"Block {KnBlockName} has no editable keyed-note attribute.");
        }

        Table table = BuildKeyNoteTable(
          database,
          plan,
          insertionPoint,
          firstNumber,
          blockId,
          attributeDefinitionId);
        BlockTableRecord space = (BlockTableRecord)transaction.GetObject(
          database.CurrentSpaceId,
          OpenMode.ForWrite);
        ObjectId tableId = space.AppendEntity(table);
        transaction.AddNewlyCreatedDBObject(table, true);
        table.GenerateLayout();

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

    private static Table BuildKeyNoteTable(
      Database database,
      KeyNoteTablePlan plan,
      Point3d insertionPoint,
      int firstNumber,
      ObjectId blockId,
      ObjectId attributeDefinitionId)
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
        if (!string.IsNullOrWhiteSpace(plan.Layer))
        {
          stage = $"assigning table layer '{plan.Layer}'";
          table.Layer = plan.Layer;
        }

        stage = "sizing table columns";
        double[] keyWidths = plan.Columns
          .Select(GetKeyNoteSymbolColumnWidth)
          .ToArray();
        for (int columnIndex = 0; columnIndex < plan.Columns.Count; columnIndex++)
        {
          KeyNoteSourceColumn column = plan.Columns[columnIndex];
          int symbolColumn = columnIndex * 2;
          int textColumn = symbolColumn + 1;
          table.Columns[symbolColumn].Width = keyWidths[columnIndex];
          double scale = GetKeyNoteColumnScale(column);
          double margin = KeyNoteTableBaseMargin * scale;
          double contentWidth = Math.Max(
            column.SourceWidth + margin * 2.0,
            column.Entries.Max(entry => entry.TextHeight) * 4.0);
          if (columnIndex + 1 < plan.Columns.Count)
          {
            double sourceDelta = plan.Columns[columnIndex + 1].SourceLeft -
              column.SourceLeft;
            double alignmentWidth = sourceDelta - keyWidths[columnIndex + 1];
            contentWidth = Math.Max(contentWidth, alignmentWidth);
          }
          else
          {
            contentWidth += KeyNoteTableColumnGap * scale;
          }
          table.Columns[textColumn].Width = contentWidth;
        }

        double[] desiredRowHeights = new double[rowCount];
        for (int row = 0; row < rowCount; row++)
        {
          double rowHeight = 0.0;
          foreach (KeyNoteSourceColumn column in plan.Columns)
          {
            if (row >= column.Entries.Count)
            {
              continue;
            }
            KeyNoteEntry entry = column.Entries[row];
            double scale = GetKeyNoteEntryScale(entry);
            double margin = KeyNoteTableBaseMargin * scale;
            rowHeight = Math.Max(
              rowHeight,
              Math.Max(
                KeyNoteHexHeight * scale + margin * 2.0,
                entry.TextHeight * 1.5 + margin * 2.0));
          }
          desiredRowHeights[row] = Math.Max(rowHeight, database.Textsize * 1.5);
        }

        int number = firstNumber;
        for (int sourceColumn = 0; sourceColumn < plan.Columns.Count; sourceColumn++)
        {
          KeyNoteSourceColumn column = plan.Columns[sourceColumn];
          int symbolColumn = sourceColumn * 2;
          int textColumn = symbolColumn + 1;
          for (int row = 0; row < rowCount; row++)
          {
            stage = $"formatting row {row + 1}, column {sourceColumn + 1}";
            HideKeyNoteCellGrid(table, row, symbolColumn);
            HideKeyNoteCellGrid(table, row, textColumn);
            EnsureKeyNoteTableCellContent(table, row, symbolColumn);
            EnsureKeyNoteTableCellContent(table, row, textColumn);
            if (row >= column.Entries.Count)
            {
              table.SetTextString(row, symbolColumn, 0, string.Empty);
              table.SetTextString(row, textColumn, 0, string.Empty);
              continue;
            }

            KeyNoteEntry entry = column.Entries[row];
            double scale = GetKeyNoteEntryScale(entry);
            double margin = KeyNoteTableBaseMargin * scale;
            SetKeyNoteCellMargins(table, row, symbolColumn, margin, margin);
            SetKeyNoteCellMargins(table, row, textColumn, margin, margin);

            stage = $"placing keyed-note symbol {number}";
            table.SetBlockTableRecordId(
              row,
              symbolColumn,
              0,
              blockId,
              false);
            table.SetIsAutoScale(row, symbolColumn, 0, false);
            table.SetScale(row, symbolColumn, 0, scale);
            table.SetBlockAttributeValue(
              row,
              symbolColumn,
              0,
              attributeDefinitionId,
              number.ToString(CultureInfo.InvariantCulture));
            table.Cells[row, symbolColumn].Alignment = CellAlignment.TopCenter;

            stage = $"writing note text for symbol {number}";
            table.SetTextString(
              row,
              textColumn,
              0,
              ConvertPlainTextToMTextContents(entry.PlainText));
            table.SetTextHeight(row, textColumn, 0, entry.TextHeight);
            table.SetTextStyleId(row, textColumn, 0, entry.TextStyleId);
            table.Cells[row, textColumn].Alignment = CellAlignment.TopLeft;
            table.SetContentColor(
              row,
              textColumn,
              0,
              ResolveKeyNoteTableColor(entry.ColorIndex));
            number++;
          }
        }

        stage = "generating the initial table layout";
        table.GenerateLayout();
        stage = "sizing table rows";
        for (int row = 0; row < rowCount; row++)
        {
          double safeHeight = Math.Max(
            desiredRowHeights[row],
            table.Rows[row].Height);
          try
          {
            safeHeight = Math.Max(safeHeight, table.Rows[row].MinimumHeight);
          }
          catch
          {
            // Some older table styles do not expose MinimumHeight until the
            // table is database-resident. Its current height is still valid.
          }
          table.Rows[row].Height = safeHeight;
        }
        stage = "generating the final table layout";
        table.GenerateLayout();
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

    private static void EnsureKeyNoteTableCellContent(
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

    private static Color ResolveKeyNoteTableColor(short colorIndex)
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

    private static void HideKeyNoteCellGrid(Table table, int row, int column)
    {
      try
      {
        table.SetGridVisibility(
          row,
          column,
          GridLineType.AllGridLines,
          Visibility.Invisible);
        return;
      }
      catch (Autodesk.AutoCAD.Runtime.Exception)
      {
        // Some older table styles reject the combined AllGridLines mask.
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

    private static void SetKeyNoteCellMargins(
      Table table,
      int row,
      int column,
      double verticalMargin,
      double horizontalMargin)
    {
      table.SetMargin(row, column, CellMargins.Top, verticalMargin);
      table.SetMargin(row, column, CellMargins.Bottom, verticalMargin);
      table.SetMargin(row, column, CellMargins.Left, horizontalMargin);
      table.SetMargin(row, column, CellMargins.Right, horizontalMargin);
    }

    private static ObjectId FindKeyNoteAttributeDefinition(
      Transaction transaction,
      ObjectId blockId)
    {
      BlockTableRecord block = transaction.GetObject(
        blockId,
        OpenMode.ForRead) as BlockTableRecord;
      ObjectId fallback = ObjectId.Null;
      foreach (ObjectId entityId in block)
      {
        AttributeDefinition definition = transaction.GetObject(
          entityId,
          OpenMode.ForRead,
          false) as AttributeDefinition;
        if (definition == null || definition.Constant)
        {
          continue;
        }
        if (fallback.IsNull)
        {
          fallback = entityId;
        }
        if (string.Equals(
          definition.Tag,
          KnAttributeTag,
          StringComparison.OrdinalIgnoreCase))
        {
          return entityId;
        }
      }
      return fallback;
    }

    private static Point3d GetKeyNoteTableDefaultPosition(KeyNoteTablePlan plan)
    {
      KeyNoteSourceColumn first = plan.Columns[0];
      double left = first.SourceLeft - GetKeyNoteSymbolColumnWidth(first);
      double top = plan.Columns.Max(column => column.SourceTop);
      double cosine = Math.Cos(plan.Rotation);
      double sine = Math.Sin(plan.Rotation);
      return new Point3d(
        left * cosine - top * sine,
        left * sine + top * cosine,
        plan.Elevation);
    }

    private static double GetKeyNoteSymbolColumnWidth(KeyNoteSourceColumn column)
    {
      double scale = GetKeyNoteColumnScale(column);
      return (KeyNoteHexWidth + KeyNoteTableBaseMargin * 2.0) * scale;
    }

    private static double GetKeyNoteColumnScale(KeyNoteSourceColumn column)
    {
      return column.Entries.Count == 0
        ? 1.0
        : column.Entries.Max(GetKeyNoteEntryScale);
    }

    private static double GetKeyNoteEntryScale(KeyNoteEntry entry)
    {
      return Math.Max(entry.TextHeight / KnAttributeHeight, 1e-6);
    }

    private static void ValidateKeyNoteRotations(
      IEnumerable<double> rotations,
      double reference)
    {
      if (rotations.Any(rotation =>
        AngularDifference(rotation, reference) > KeyNoteRotationTolerance))
      {
        throw new InvalidOperationException(
          "Selected text objects must share the same rotation (within one degree).");
      }
    }

    private static double AngularDifference(double first, double second)
    {
      double difference = Math.Abs(first - second) % (Math.PI * 2.0);
      return difference > Math.PI
        ? Math.PI * 2.0 - difference
        : difference;
    }

    private static LocalBounds GetLocalBounds(Entity entity, double rotation)
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
      return new LocalBounds
      {
        Left = localX.Min(),
        Right = localX.Max(),
        Bottom = localY.Min(),
        Top = localY.Max(),
      };
    }

    private static double PositiveOrDefault(double value, double fallback)
    {
      if (value > 1e-9)
      {
        return value;
      }
      return fallback > 1e-9 ? fallback : KnAttributeHeight;
    }

    private static double Median(IEnumerable<double> values)
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

    private sealed class KeyNoteTablePlan
    {
      internal List<KeyNoteSourceColumn> Columns { get; set; } =
        new List<KeyNoteSourceColumn>();
      internal double Rotation { get; set; }
      internal double Elevation { get; set; }
      internal string Layer { get; set; } = string.Empty;
      internal int DbVisualLineCount { get; set; }
      internal int NoteCount => Columns.Sum(column => column.Entries.Count);
    }

    private sealed class KeyNoteSourceColumn
    {
      internal List<KeyNoteEntry> Entries { get; } = new List<KeyNoteEntry>();
      internal double SourceLeft { get; set; }
      internal double SourceTop { get; set; }
      internal double SourceWidth { get; set; }
      internal int DbVisualLineCount { get; set; }
    }

    private sealed class KeyNoteEntry
    {
      internal string PlainText { get; set; } = string.Empty;
      internal double TextHeight { get; set; }
      internal ObjectId TextStyleId { get; set; }
      internal short ColorIndex { get; set; } = 256;
    }

    private sealed class DbTextLine
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

    private sealed class DbVisualLine
    {
      internal List<DbTextLine> Parts { get; } = new List<DbTextLine>();
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
        List<DbTextLine> ordered = Parts.OrderBy(part => part.Left).ToList();
        Text = string.Join(" ", ordered.Select(part => part.Text));
        Left = ordered.Min(part => part.Left);
        Right = ordered.Max(part => part.Right);
        Top = ordered.Max(part => part.Top);
        CenterY = Median(ordered.Select(part => part.CenterY));
        TextHeight = ordered.Max(part => part.TextHeight);
        TextStyleId = ordered[0].TextStyleId;
        ColorIndex = ordered[0].ColorIndex;
      }
    }

    private sealed class LocalBounds
    {
      internal double Left { get; set; }
      internal double Right { get; set; }
      internal double Bottom { get; set; }
      internal double Top { get; set; }
      internal double Width => Right - Left;
    }
  }
}
