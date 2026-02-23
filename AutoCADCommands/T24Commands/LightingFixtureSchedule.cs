using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string LightingFixtureScheduleTemplateResourceSuffix = "Resources.LightingFixtureSchedule.tableanalyze.json";

    [CommandMethod("LIGHTINGFIXTURESCHEDULE", CommandFlags.Modal)]
    [CommandMethod("LFS", CommandFlags.Modal)]
    public void LightingFixtureSchedule()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      var warnings = new List<LightingFixtureScheduleWarning>();
      string resourceName = string.Empty;

      try
      {
        TableAnalyzeExport template = LoadTemplateFromEmbeddedResource(out resourceName);
        ValidateTemplate(template, warnings);

        PromptPointResult pointResult = ed.GetPoint("\nSelect insertion point for lighting fixture schedule: ");
        if (pointResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nInsertion cancelled.");
          return;
        }

        ObjectId newTableId = ObjectId.Null;
        string handle = "UNKNOWN";
        bool syncApplied = false;
        string syncMessage = string.Empty;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          Table table = RecreateTable(template, db, tr, pointResult.Value, warnings);

          BlockTableRecord currentSpace = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          if (currentSpace == null)
          {
            throw new InvalidOperationException("Unable to open current space for table insertion.");
          }

          newTableId = currentSpace.AppendEntity(table);
          tr.AddNewlyCreatedDBObject(table, true);

          bool syncError;
          if (TryApplyLightingFixtureScheduleSyncToTable(table, db, out syncMessage, out syncError))
          {
            syncApplied = true;
          }
          else if (syncError && !string.IsNullOrWhiteSpace(syncMessage))
          {
            AddWarning(warnings, "sync.apply", syncMessage);
          }

          SafeApply(
            warnings,
            "table.generateLayout",
            () => table.GenerateLayout()
          );
          SafeApply(
            warnings,
            "table.recomputeTableBlock",
            () => TryRecomputeTableBlock(table)
          );

          handle = SafeGet(
            warnings,
            "table.handle",
            () => table.Handle.ToString(),
            null,
            null,
            "UNKNOWN"
          );

          tr.Commit();
        }

        ed.WriteMessage(
          $"\nLIGHTINGFIXTURESCHEDULE complete. Handle: {handle}, ObjectId: {newTableId}, Rows: {template.Table.NumRows}, Columns: {template.Table.NumColumns}, Warnings: {warnings.Count}, Template: {resourceName}, SyncApplied: {syncApplied}"
        );
        if (syncApplied && !string.IsNullOrWhiteSpace(syncMessage))
        {
          ed.WriteMessage($"\n{syncMessage}");
        }

        if (warnings.Count > 0)
        {
          foreach (LightingFixtureScheduleWarning warning in warnings.Take(10))
          {
            string location = (warning.Row.HasValue && warning.Column.HasValue)
              ? $" row {warning.Row.Value}, col {warning.Column.Value}"
              : string.Empty;
            ed.WriteMessage($"\nWarning{location} [{warning.Property}]: {warning.Message}");
          }

          if (warnings.Count > 10)
          {
            ed.WriteMessage($"\n... {warnings.Count - 10} additional warnings omitted.");
          }
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nLIGHTINGFIXTURESCHEDULE error: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    private static TableAnalyzeExport LoadTemplateFromEmbeddedResource(out string resourceName)
    {
      Assembly assembly = typeof(GeneralCommands).Assembly;
      resourceName = assembly
        .GetManifestResourceNames()
        .FirstOrDefault(name =>
          name.EndsWith(LightingFixtureScheduleTemplateResourceSuffix, StringComparison.OrdinalIgnoreCase)
        );

      if (string.IsNullOrWhiteSpace(resourceName))
      {
        throw new FileNotFoundException(
          $"Embedded template resource ending with '{LightingFixtureScheduleTemplateResourceSuffix}' was not found."
        );
      }

      using (Stream stream = assembly.GetManifestResourceStream(resourceName))
      {
        if (stream == null)
        {
          throw new FileNotFoundException($"Unable to open embedded template resource '{resourceName}'.");
        }

        using (var reader = new StreamReader(stream))
        {
          string json = reader.ReadToEnd();
          TableAnalyzeExport export = JsonConvert.DeserializeObject<TableAnalyzeExport>(json);
          if (export == null)
          {
            throw new InvalidDataException($"Failed to deserialize template resource '{resourceName}'.");
          }

          return export;
        }
      }
    }

    private static void ValidateTemplate(TableAnalyzeExport export, List<LightingFixtureScheduleWarning> warnings)
    {
      if (export == null)
      {
        throw new InvalidDataException("Template payload is null.");
      }

      if (export.Table == null)
      {
        throw new InvalidDataException("Template table section is missing.");
      }

      if (export.Table.NumRows <= 0 || export.Table.NumColumns <= 0)
      {
        throw new InvalidDataException("Template table dimensions are invalid.");
      }

      if (!string.Equals(export.SchemaVersion, "1.0.0", StringComparison.OrdinalIgnoreCase))
      {
        AddWarning(
          warnings,
          "template.schemaVersion",
          $"Expected schemaVersion '1.0.0' but found '{export.SchemaVersion ?? "(null)"}'. Continuing with best effort."
        );
      }

      int expectedCellCount = export.Table.NumRows * export.Table.NumColumns;
      int actualCellCount = export.Cells?.Count ?? 0;
      if (actualCellCount != expectedCellCount)
      {
        AddWarning(
          warnings,
          "template.cells",
          $"Cell count mismatch. Expected {expectedCellCount}, actual {actualCellCount}. Continuing with best effort."
        );
      }

      export.Cells = export.Cells ?? new List<TableAnalyzeCell>();
      export.MergedRanges = export.MergedRanges ?? new List<TableAnalyzeMergedRange>();
    }

    private static Table RecreateTable(
      TableAnalyzeExport export,
      Database db,
      Transaction tr,
      Point3d insertionPoint,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      var table = new Table();
      table.SetDatabaseDefaults();

      ObjectId tableStyleId = ResolveTableStyleId(db, tr, export.Table.TableStyleName, warnings);
      SafeApply(warnings, "table.tableStyle", () => table.TableStyle = tableStyleId);

      SafeApply(warnings, "table.titleSuppressed", () => table.IsTitleSuppressed = export.Table.IsTitleSuppressed);
      SafeApply(warnings, "table.headerSuppressed", () => table.IsHeaderSuppressed = export.Table.IsHeaderSuppressed);
      SafeApply(warnings, "table.breakEnabled", () => table.BreakEnabled = export.Table.BreakEnabled);

      FlowDirection flowDirection = ParseEnum(
        export.Table.FlowDirection,
        FlowDirection.TopToBottom,
        warnings,
        "table.flowDirection"
      );
      SafeApply(warnings, "table.flowDirection", () => table.FlowDirection = flowDirection);

      SafeApply(warnings, "table.setSize", () => table.SetSize(export.Table.NumRows, export.Table.NumColumns));

      ApplyRowHeights(table, export.Table, warnings);
      ApplyColumnWidths(table, export.Table, warnings);

      SafeApply(warnings, "table.horizontalCellMargin", () => table.HorizontalCellMargin = export.Table.HorizontalCellMargin);
      SafeApply(warnings, "table.verticalCellMargin", () => table.VerticalCellMargin = export.Table.VerticalCellMargin);

      if (export.Table.Direction != null)
      {
        Vector3d direction = new Vector3d(export.Table.Direction.X, export.Table.Direction.Y, export.Table.Direction.Z);
        SafeApply(warnings, "table.direction", () => table.Direction = direction);
      }

      if (export.Table.Normal != null)
      {
        Vector3d normal = new Vector3d(export.Table.Normal.X, export.Table.Normal.Y, export.Table.Normal.Z);
        SafeApply(warnings, "table.normal", () => table.Normal = normal);
      }

      SafeApply(warnings, "table.rotation", () => table.Rotation = export.Table.Rotation);

      string layerName = ResolveLayerName(db, tr, export.Source?.Layer, warnings);
      SafeApply(warnings, "table.layer", () => table.Layer = layerName);

      SafeApply(warnings, "table.position", () => table.Position = insertionPoint);

      ApplyMergedRanges(table, export.MergedRanges, warnings);
      ApplyCellData(table, db, tr, export, warnings);

      return table;
    }

    private static void ApplyRowHeights(Table table, TableAnalyzeTable tableData, List<LightingFixtureScheduleWarning> warnings)
    {
      for (int row = 0; row < tableData.NumRows; row++)
      {
        double rowHeight = row < tableData.RowHeights.Count ? tableData.RowHeights[row] : 0.2;
        SafeApply(warnings, "table.rowHeight", () => table.SetRowHeight(row, rowHeight), row, null);
      }
    }

    private static void ApplyColumnWidths(Table table, TableAnalyzeTable tableData, List<LightingFixtureScheduleWarning> warnings)
    {
      for (int col = 0; col < tableData.NumColumns; col++)
      {
        double columnWidth = col < tableData.ColumnWidths.Count ? tableData.ColumnWidths[col] : 0.5;
        SafeApply(warnings, "table.columnWidth", () => table.SetColumnWidth(col, columnWidth), null, col);
      }
    }

    private static void ApplyMergedRanges(Table table, List<TableAnalyzeMergedRange> mergedRanges, List<LightingFixtureScheduleWarning> warnings)
    {
      foreach (TableAnalyzeMergedRange range in mergedRanges)
      {
        if (range == null)
        {
          continue;
        }

        SafeApply(
          warnings,
          "table.mergeCells",
          () => table.MergeCells(CellRange.Create(table, range.TopRow, range.LeftColumn, range.BottomRow, range.RightColumn)),
          range.TopRow,
          range.LeftColumn
        );
      }
    }

    private static void ApplyCellData(
      Table table,
      Database db,
      Transaction tr,
      TableAnalyzeExport export,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      var cellMap = export.Cells
        .Where(cell => cell != null)
        .GroupBy(cell => Tuple.Create(cell.Row, cell.Column))
        .ToDictionary(group => group.Key, group => group.First());

      for (int row = 0; row < export.Table.NumRows; row++)
      {
        for (int col = 0; col < export.Table.NumColumns; col++)
        {
          TableAnalyzeCell cell;
          if (!cellMap.TryGetValue(Tuple.Create(row, col), out cell) || cell == null)
          {
            continue;
          }

          ApplyCellProperties(table, db, tr, row, col, cell, warnings);
          ApplyGridLines(table, db, tr, row, col, cell, warnings);
          ApplyCellContents(table, db, tr, row, col, cell, warnings);
        }
      }
    }

    private static void ApplyCellProperties(
      Table table,
      Database db,
      Transaction tr,
      int row,
      int col,
      TableAnalyzeCell cell,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (!string.IsNullOrWhiteSpace(cell.CellStyle))
      {
        SafeApply(warnings, "cell.cellStyle", () => table.SetCellStyle(row, col, cell.CellStyle), row, col);
      }

      if (!string.IsNullOrWhiteSpace(cell.ToolTip))
      {
        SafeApply(warnings, "cell.toolTip", () => table.SetToolTip(row, col, cell.ToolTip), row, col);
      }

      CellAlignment alignment = ParseEnum(cell.Alignment, CellAlignment.MiddleLeft, warnings, "cell.alignment", row, col);
      SafeApply(warnings, "cell.alignment", () => table.SetAlignment(row, col, alignment), row, col);

      RotationAngle textRotation = ParseEnum(cell.TextRotation, RotationAngle.Degrees000, warnings, "cell.textRotation", row, col);
      SafeApply(warnings, "cell.textRotation", () => table.SetTextRotation(row, col, textRotation), row, col);

      CellContentLayout contentLayout = ParseEnum(cell.ContentLayout, CellContentLayout.Flow, warnings, "cell.contentLayout", row, col);
      SafeApply(warnings, "cell.contentLayout", () => table.SetContentLayout(row, col, contentLayout), row, col);

      DataType dataType = ParseEnum(cell.DataType, DataType.String, warnings, "cell.dataType", row, col);
      UnitType unitType = ParseEnum(cell.UnitType, UnitType.Unitless, warnings, "cell.unitType", row, col);
      SafeApply(warnings, "cell.dataType", () => table.SetDataType(row, col, dataType, unitType), row, col);

      if (!string.IsNullOrWhiteSpace(cell.DataFormat))
      {
        SafeApply(warnings, "cell.dataFormat", () => table.SetDataFormat(row, col, cell.DataFormat), row, col);
      }

      if (cell.TopMargin.HasValue)
      {
        SafeApply(warnings, "cell.marginTop", () => table.SetMargin(row, col, CellMargins.Top, cell.TopMargin.Value), row, col);
      }
      if (cell.LeftMargin.HasValue)
      {
        SafeApply(warnings, "cell.marginLeft", () => table.SetMargin(row, col, CellMargins.Left, cell.LeftMargin.Value), row, col);
      }
      if (cell.BottomMargin.HasValue)
      {
        SafeApply(warnings, "cell.marginBottom", () => table.SetMargin(row, col, CellMargins.Bottom, cell.BottomMargin.Value), row, col);
      }
      if (cell.RightMargin.HasValue)
      {
        SafeApply(warnings, "cell.marginRight", () => table.SetMargin(row, col, CellMargins.Right, cell.RightMargin.Value), row, col);
      }

      if (cell.TextHeight.HasValue)
      {
        SafeApply(warnings, "cell.textHeight", () => table.SetTextHeight(row, col, cell.TextHeight.Value), row, col);
      }

      ObjectId textStyleId = ResolveTextStyleId(db, tr, cell.TextStyle?.Name, warnings, row, col);
      if (!textStyleId.IsNull)
      {
        SafeApply(warnings, "cell.textStyle", () => table.SetTextStyle(row, col, textStyleId), row, col);
      }

      Color contentColor = ToAcadColor(cell.ContentColor, warnings, "cell.contentColor", row, col);
      if (contentColor != null)
      {
        SafeApply(warnings, "cell.contentColor", () => table.SetContentColor(row, col, contentColor), row, col);
      }

      if (cell.IsBackgroundColorNone.HasValue)
      {
        SafeApply(warnings, "cell.backgroundColorNone", () => table.SetBackgroundColorNone(row, col, cell.IsBackgroundColorNone.Value), row, col);
      }

      if (cell.BackgroundColor != null && !(cell.IsBackgroundColorNone ?? false))
      {
        Color backgroundColor = ToAcadColor(cell.BackgroundColor, warnings, "cell.backgroundColor", row, col);
        if (backgroundColor != null)
        {
          SafeApply(warnings, "cell.backgroundColor", () => table.SetBackgroundColor(row, col, backgroundColor), row, col);
        }
      }
    }

    private static void ApplyGridLines(
      Table table,
      Database db,
      Transaction tr,
      int row,
      int col,
      TableAnalyzeCell cell,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (cell.GridLines == null || cell.GridLines.Count == 0)
      {
        return;
      }

      foreach (TableAnalyzeGridLine grid in cell.GridLines)
      {
        if (grid == null)
        {
          continue;
        }

        GridLineType gridType = ParseEnum(grid.GridLineType, GridLineType.AllGridLines, warnings, "grid.type", row, col);
        Visibility visibility = ParseEnum(grid.Visibility, Visibility.Visible, warnings, "grid.visibility", row, col);
        GridLineStyle lineStyle = ParseEnum(grid.LineStyle, GridLineStyle.Single, warnings, "grid.lineStyle", row, col);
        LineWeight lineWeight = ParseEnum(grid.LineWeight, LineWeight.ByBlock, warnings, "grid.lineWeight", row, col);

        SafeApply(warnings, "grid.visibility", () => table.SetGridVisibility(row, col, gridType, visibility), row, col);
        SafeApply(warnings, "grid.lineStyle", () => table.SetGridLineStyle(row, col, gridType, lineStyle), row, col);
        SafeApply(warnings, "grid.lineWeight", () => table.SetGridLineWeight(row, col, gridType, lineWeight), row, col);

        ObjectId linetypeId = ResolveLinetypeId(db, tr, grid.Linetype?.Name, warnings, row, col);
        if (!linetypeId.IsNull)
        {
          SafeApply(warnings, "grid.linetype", () => table.SetGridLinetype(row, col, gridType, linetypeId), row, col);
        }

        Color color = ToAcadColor(grid.Color, warnings, "grid.color", row, col);
        if (color != null)
        {
          SafeApply(warnings, "grid.color", () => table.SetGridColor(row, col, gridType, color), row, col);
        }

        if (grid.DoubleLineSpacing.HasValue)
        {
          SafeApply(
            warnings,
            "grid.doubleLineSpacing",
            () => table.SetGridDoubleLineSpacing(row, col, gridType, grid.DoubleLineSpacing.Value),
            row,
            col
          );
        }
      }
    }

    private static void ApplyCellContents(
      Table table,
      Database db,
      Transaction tr,
      int row,
      int col,
      TableAnalyzeCell cell,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (cell.Contents == null || cell.Contents.Count == 0)
      {
        return;
      }

      List<TableAnalyzeCellContent> orderedContents = cell.Contents
        .Where(content => content != null)
        .OrderBy(content => content.ContentIndex)
        .ToList();

      int requiredCount = orderedContents.Count;
      int currentCount = SafeGet(warnings, "content.currentCount", () => table.GetNumberOfContents(row, col), row, col, 0);

      for (int contentIndex = currentCount; contentIndex < requiredCount; contentIndex++)
      {
        int idx = contentIndex;
        SafeApply(warnings, "content.create", () => table.CreateContent(row, col, idx), row, col);
      }

      foreach (TableAnalyzeCellContent content in orderedContents)
      {
        int contentIndex = content.ContentIndex;
        if (contentIndex < 0)
        {
          AddWarning(warnings, "content.index", "Negative content index ignored.", row, col);
          continue;
        }

        string text = GetPreferredContentText(content);
        if (!string.IsNullOrEmpty(text))
        {
          SafeApply(warnings, "content.textString", () => table.SetTextString(row, col, contentIndex, text), row, col);
        }

        DataType dataType = ParseEnum(content.DataType, DataType.String, warnings, "content.dataType", row, col);
        UnitType unitType = ParseEnum(content.UnitType, UnitType.Unitless, warnings, "content.unitType", row, col);
        var dataTypeParameter = new DataTypeParameter(dataType, unitType);
        SafeApply(warnings, "content.dataType", () => table.SetDataType(row, col, contentIndex, dataTypeParameter), row, col);

        if (!string.IsNullOrWhiteSpace(content.DataFormat))
        {
          SafeApply(warnings, "content.dataFormat", () => table.SetDataFormat(row, col, contentIndex, content.DataFormat), row, col);
        }

        if (content.TextHeight.HasValue)
        {
          SafeApply(warnings, "content.textHeight", () => table.SetTextHeight(row, col, contentIndex, content.TextHeight.Value), row, col);
        }

        ObjectId contentTextStyleId = ResolveTextStyleId(db, tr, content.TextStyle?.Name, warnings, row, col);
        if (!contentTextStyleId.IsNull)
        {
          SafeApply(warnings, "content.textStyle", () => table.SetTextStyleId(row, col, contentIndex, contentTextStyleId), row, col);
        }

        Color contentColor = ToAcadColor(content.ContentColor, warnings, "content.contentColor", row, col);
        if (contentColor != null)
        {
          SafeApply(warnings, "content.contentColor", () => table.SetContentColor(row, col, contentIndex, contentColor), row, col);
        }

        if (content.IsAutoScale.HasValue)
        {
          SafeApply(warnings, "content.autoScale", () => table.SetIsAutoScale(row, col, contentIndex, content.IsAutoScale.Value), row, col);
        }

        if (content.Rotation.HasValue)
        {
          SafeApply(warnings, "content.rotation", () => table.SetRotation(row, col, contentIndex, content.Rotation.Value), row, col);
        }

        if (content.Scale.HasValue)
        {
          SafeApply(warnings, "content.scale", () => table.SetScale(row, col, contentIndex, content.Scale.Value), row, col);
        }

        if ((content.HasFormula ?? false) && !string.IsNullOrWhiteSpace(content.Formula))
        {
          SafeApply(warnings, "content.formula", () => table.SetFormula(row, col, contentIndex, content.Formula), row, col);
        }
      }
    }

    private static string GetPreferredContentText(TableAnalyzeCellContent content)
    {
      if (!string.IsNullOrEmpty(content.TextString)) return content.TextString;
      if (!string.IsNullOrEmpty(content.TextStringForEditing)) return content.TextStringForEditing;
      if (!string.IsNullOrEmpty(content.TextStringFromContentObject)) return content.TextStringFromContentObject;
      if (!string.IsNullOrEmpty(content.TextStringForExpression)) return content.TextStringForExpression;
      return content.TextStringWithoutMtext ?? string.Empty;
    }

    private static string ResolveLayerName(
      Database db,
      Transaction tr,
      string preferredLayer,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
      if (layerTable != null && !string.IsNullOrWhiteSpace(preferredLayer) && layerTable.Has(preferredLayer))
      {
        return preferredLayer;
      }

      if (!string.IsNullOrWhiteSpace(preferredLayer))
      {
        AddWarning(warnings, "table.layer", $"Layer '{preferredLayer}' not found. Falling back to current layer.");
      }

      LayerTableRecord currentLayer = tr.GetObject(db.Clayer, OpenMode.ForRead) as LayerTableRecord;
      return currentLayer?.Name ?? "0";
    }

    private static ObjectId ResolveTableStyleId(
      Database db,
      Transaction tr,
      string styleName,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      try
      {
        DBDictionary styleDict = tr.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead) as DBDictionary;
        if (styleDict != null && !string.IsNullOrWhiteSpace(styleName))
        {
          if (styleDict.Contains(styleName))
          {
            return styleDict.GetAt(styleName);
          }

          foreach (DBDictionaryEntry entry in styleDict)
          {
            if (string.Equals(entry.Key, styleName, StringComparison.OrdinalIgnoreCase))
            {
              return entry.Value;
            }
          }

          AddWarning(warnings, "table.tableStyle", $"Table style '{styleName}' was not found. Falling back to current database table style.");
        }
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, "table.tableStyle", ex.Message);
      }

      return db.Tablestyle;
    }

    private static ObjectId ResolveTextStyleId(
      Database db,
      Transaction tr,
      string styleName,
      List<LightingFixtureScheduleWarning> warnings,
      int? row,
      int? col
    )
    {
      if (string.IsNullOrWhiteSpace(styleName))
      {
        return db.Textstyle;
      }

      TextStyleTable textStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
      if (textStyleTable == null)
      {
        AddWarning(warnings, "textStyle.resolve", "Text style table is unavailable. Using current text style.", row, col);
        return db.Textstyle;
      }

      if (textStyleTable.Has(styleName))
      {
        return textStyleTable[styleName];
      }

      foreach (ObjectId id in textStyleTable)
      {
        TextStyleTableRecord record = tr.GetObject(id, OpenMode.ForRead) as TextStyleTableRecord;
        if (record != null && string.Equals(record.Name, styleName, StringComparison.OrdinalIgnoreCase))
        {
          return id;
        }
      }

      AddWarning(warnings, "textStyle.resolve", $"Text style '{styleName}' not found. Using current text style.", row, col);
      return db.Textstyle;
    }

    private static ObjectId ResolveLinetypeId(
      Database db,
      Transaction tr,
      string linetypeName,
      List<LightingFixtureScheduleWarning> warnings,
      int? row,
      int? col
    )
    {
      if (string.IsNullOrWhiteSpace(linetypeName))
      {
        return db.Celtype;
      }

      LinetypeTable linetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
      if (linetypeTable == null)
      {
        AddWarning(warnings, "linetype.resolve", "Linetype table is unavailable. Using current linetype.", row, col);
        return db.Celtype;
      }

      if (linetypeTable.Has(linetypeName))
      {
        return linetypeTable[linetypeName];
      }

      foreach (ObjectId id in linetypeTable)
      {
        LinetypeTableRecord record = tr.GetObject(id, OpenMode.ForRead) as LinetypeTableRecord;
        if (record != null && string.Equals(record.Name, linetypeName, StringComparison.OrdinalIgnoreCase))
        {
          return id;
        }
      }

      AddWarning(warnings, "linetype.resolve", $"Linetype '{linetypeName}' not found. Using current linetype.", row, col);
      return db.Celtype;
    }

    private static Color ToAcadColor(
      TableAnalyzeColor color,
      List<LightingFixtureScheduleWarning> warnings,
      string property,
      int? row,
      int? col
    )
    {
      if (color == null)
      {
        return null;
      }

      try
      {
        string method = color.ColorMethod ?? string.Empty;
        if (method.Equals("ByLayer", StringComparison.OrdinalIgnoreCase))
        {
          return Color.FromColorIndex(ColorMethod.ByLayer, 256);
        }

        if (method.Equals("ByBlock", StringComparison.OrdinalIgnoreCase))
        {
          return Color.FromColorIndex(ColorMethod.ByBlock, 0);
        }

        if (method.Equals("ByAci", StringComparison.OrdinalIgnoreCase))
        {
          short aci = color.ColorIndex;
          if (aci < 0)
          {
            aci = 7;
          }

          return Color.FromColorIndex(ColorMethod.ByAci, aci);
        }

        if (method.Equals("None", StringComparison.OrdinalIgnoreCase) || color.IsNone)
        {
          return Color.FromColorIndex(ColorMethod.None, 257);
        }

        if (method.Equals("ByColor", StringComparison.OrdinalIgnoreCase) || method.Equals("ByRgb", StringComparison.OrdinalIgnoreCase))
        {
          return Color.FromRgb(color.Red, color.Green, color.Blue);
        }

        if (color.Red != 0 || color.Green != 0 || color.Blue != 0)
        {
          return Color.FromRgb(color.Red, color.Green, color.Blue);
        }

        if (color.ColorIndex >= 0)
        {
          return Color.FromColorIndex(ColorMethod.ByAci, color.ColorIndex);
        }
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, property, ex.Message, row, col);
      }

      AddWarning(warnings, property, "Unable to convert captured color. Falling back to ByLayer.", row, col);
      return Color.FromColorIndex(ColorMethod.ByLayer, 256);
    }

    private static TEnum ParseEnum<TEnum>(
      string value,
      TEnum fallback,
      List<LightingFixtureScheduleWarning> warnings,
      string property,
      int? row = null,
      int? col = null
    ) where TEnum : struct
    {
      if (string.IsNullOrWhiteSpace(value))
      {
        return fallback;
      }

      TEnum parsed;
      if (Enum.TryParse(value, true, out parsed))
      {
        return parsed;
      }

      AddWarning(
        warnings,
        property,
        $"Invalid {typeof(TEnum).Name} value '{value}'. Using '{fallback}'.",
        row,
        col
      );
      return fallback;
    }

    private static void SafeApply(
      List<LightingFixtureScheduleWarning> warnings,
      string property,
      Action action,
      int? row = null,
      int? col = null
    )
    {
      try
      {
        action();
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, property, ex.Message, row, col);
      }
    }

    private static T SafeGet<T>(
      List<LightingFixtureScheduleWarning> warnings,
      string property,
      Func<T> getter,
      int? row,
      int? col,
      T fallback
    )
    {
      try
      {
        return getter();
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, property, ex.Message, row, col);
        return fallback;
      }
    }

    private static void AddWarning(
      List<LightingFixtureScheduleWarning> warnings,
      string property,
      string message,
      int? row = null,
      int? col = null
    )
    {
      warnings.Add(new LightingFixtureScheduleWarning
      {
        Property = property,
        Message = message,
        Row = row,
        Column = col
      });
    }
  }

  public class LightingFixtureScheduleWarning
  {
    public int? Row { get; set; }
    public int? Column { get; set; }
    public string Property { get; set; }
    public string Message { get; set; }
  }
}
