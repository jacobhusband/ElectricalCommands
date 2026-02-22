using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("TABLEANALYZE", CommandFlags.UsePickSet)]
    [CommandMethod("TBA", CommandFlags.UsePickSet)]
    public void TableAnalyze()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      try
      {
        ObjectId tableId = ResolveTableSelectionId(ed, db);
        if (tableId == ObjectId.Null)
        {
          ed.WriteMessage("\nNo table selected.");
          return;
        }

        TableAnalyzeExport export;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          Table table = tr.GetObject(tableId, OpenMode.ForRead, false) as Table;
          if (table == null)
          {
            ed.WriteMessage("\nSelected object is not an AutoCAD table.");
            return;
          }

          export = CaptureTable(table, tr, db);
          tr.Commit();
        }

        string outputPath = WriteExportToJson(export, db);
        ed.WriteMessage(
          $"\nTABLEANALYZE complete. Handle: {export.Source?.TableHandle ?? "UNKNOWN"}, Cells: {export.Cells.Count}, Merged ranges: {export.MergedRanges.Count}, Warnings: {export.Warnings.Count}, Output: {outputPath}"
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nTABLEANALYZE error: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    private static ObjectId ResolveTableSelectionId(Editor ed, Database db)
    {
      PromptSelectionResult psr = ed.SelectImplied();
      if (psr.Status == PromptStatus.OK && psr.Value != null && psr.Value.Count > 0)
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId id in psr.Value.GetObjectIds())
          {
            if (tr.GetObject(id, OpenMode.ForRead, false) is Table)
            {
              tr.Commit();
              return id;
            }
          }
          tr.Commit();
        }
      }

      PromptEntityOptions peo = new PromptEntityOptions("\nSelect an AutoCAD table to analyze: ");
      peo.SetRejectMessage("\nSelected object is not an AutoCAD table.");
      peo.AddAllowedClass(typeof(Table), true);
      PromptEntityResult per = ed.GetEntity(peo);
      return per.Status == PromptStatus.OK ? per.ObjectId : ObjectId.Null;
    }

    private static TableAnalyzeExport CaptureTable(Table table, Transaction tr, Database db)
    {
      var warnings = new List<TableAnalyzeWarning>();
      var mergedRanges = new List<TableAnalyzeMergedRange>();
      var mergedRangeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      var cells = new List<TableAnalyzeCell>();

      int rowCount = SafeGet(warnings, "table.numRows", () => table.NumRows, null, null, null, 0);
      int colCount = SafeGet(warnings, "table.numColumns", () => table.NumColumns, null, null, null, 0);

      var tableDto = new TableAnalyzeTable
      {
        TableStyle = ToObjectRef(
          SafeGet(warnings, "table.styleId", () => table.TableStyle, null, null, null, ObjectId.Null),
          tr,
          warnings,
          "table.style"
        ),
        TableStyleName = SafeGet(warnings, "table.styleName", () => table.TableStyleName, null, null, null, string.Empty),
        TableStyleOverrides = ToEnumNameList(SafeGet(warnings, "table.styleOverrides", () => table.TableStyleOverrides(), null, null, null, new TableStyleOverride[0])),
        TableStyleOverrideValues = ToEnumValueList(SafeGet(warnings, "table.styleOverrideValues", () => table.TableStyleOverrides(), null, null, null, new TableStyleOverride[0])),
        NumRows = rowCount,
        NumColumns = colCount,
        Position = ToPointDto(SafeGet(warnings, "table.position", () => table.Position, null, null, null, Point3d.Origin)),
        Direction = ToVectorDto(SafeGet(warnings, "table.direction", () => table.Direction, null, null, null, Vector3d.XAxis)),
        Normal = ToVectorDto(SafeGet(warnings, "table.normal", () => table.Normal, null, null, null, Vector3d.ZAxis)),
        Rotation = SafeGet(warnings, "table.rotation", () => table.Rotation, null, null, null, 0.0),
        Width = SafeGet(warnings, "table.width", () => table.Width, null, null, null, 0.0),
        Height = SafeGet(warnings, "table.height", () => table.Height, null, null, null, 0.0),
        MinimumTableWidth = SafeGet(warnings, "table.minWidth", () => table.MinimumTableWidth, null, null, null, 0.0),
        MinimumTableHeight = SafeGet(warnings, "table.minHeight", () => table.MinimumTableHeight, null, null, null, 0.0),
        HorizontalCellMargin = SafeGet(warnings, "table.hMargin", () => table.HorizontalCellMargin, null, null, null, 0.0),
        VerticalCellMargin = SafeGet(warnings, "table.vMargin", () => table.VerticalCellMargin, null, null, null, 0.0),
        FlowDirection = SafeGet(warnings, "table.flowDir", () => table.FlowDirection.ToString(), null, null, null, string.Empty),
        MergeStyle = SafeGet(warnings, "table.mergeStyle", () => table.MergeStyle.ToString(), null, null, null, string.Empty),
        IsTitleSuppressed = SafeGet(warnings, "table.titleSuppressed", () => table.IsTitleSuppressed, null, null, null, false),
        IsHeaderSuppressed = SafeGet(warnings, "table.headerSuppressed", () => table.IsHeaderSuppressed, null, null, null, false),
        BreakEnabled = SafeGet(warnings, "table.breakEnabled", () => table.BreakEnabled, null, null, null, false)
      };

      CaptureRowsAndColumns(table, warnings, rowCount, colCount, tableDto);
      CaptureCells(table, tr, warnings, rowCount, colCount, cells, mergedRanges, mergedRangeKeys);

      string dwgPath = db.Filename ?? string.Empty;
      var source = new TableAnalyzeSource
      {
        DwgPath = dwgPath,
        DrawingName = string.IsNullOrWhiteSpace(dwgPath) ? "(Unsaved Drawing)" : Path.GetFileName(dwgPath),
        TableHandle = SafeGet(warnings, "source.tableHandle", () => table.Handle.ToString(), null, null, null, string.Empty),
        TableObjectId = table.ObjectId.ToString(),
        Layer = SafeGet(warnings, "source.layer", () => table.Layer, null, null, null, string.Empty),
        Space = ResolveSpaceName(table, tr, warnings),
        CaptureToolVersion = typeof(GeneralCommands).Assembly.GetName().Version?.ToString() ?? "0.0.0.0"
      };

      return new TableAnalyzeExport
      {
        SchemaVersion = "1.0.0",
        CapturedAtUtc = DateTime.UtcNow,
        Source = source,
        Table = tableDto,
        MergedRanges = mergedRanges,
        Cells = cells,
        Warnings = warnings
      };
    }

    private static void CaptureRowsAndColumns(
      Table table,
      List<TableAnalyzeWarning> warnings,
      int rowCount,
      int colCount,
      TableAnalyzeTable tableDto
    )
    {
      for (int row = 0; row < rowCount; row++)
      {
        tableDto.RowHeights.Add(SafeGet(warnings, "table.rowHeight", () => table.RowHeight(row), row, null, null, 0.0));
        tableDto.RowTypes.Add(SafeGet(warnings, "table.rowType", () => table.RowType(row).ToString(), row, null, null, string.Empty));
      }

      for (int col = 0; col < colCount; col++)
      {
        tableDto.ColumnWidths.Add(SafeGet(warnings, "table.columnWidth", () => table.ColumnWidth(col), null, col, null, 0.0));
        tableDto.ColumnNames.Add(SafeGet(warnings, "table.columnName", () => table.GetColumnName(col), null, col, null, string.Empty));
      }
    }

    private static void CaptureCells(
      Table table,
      Transaction tr,
      List<TableAnalyzeWarning> warnings,
      int rowCount,
      int colCount,
      List<TableAnalyzeCell> cells,
      List<TableAnalyzeMergedRange> mergedRanges,
      HashSet<string> mergedRangeKeys
    )
    {
      for (int row = 0; row < rowCount; row++)
      {
        for (int col = 0; col < colCount; col++)
        {
          var cellDto = new TableAnalyzeCell { Row = row, Column = col };
          Cell cell = SafeGet(warnings, "cell.object", () => table.Cells[row, col], row, col, null, null as Cell);

          CaptureCellMerge(table, warnings, row, col, cellDto, mergedRanges, mergedRangeKeys);

          cellDto.RowType = SafeGet(warnings, "cell.rowType", () => table.RowType(row).ToString(), row, col, null, string.Empty);
          cellDto.CellType = SafeGet(warnings, "cell.cellType", () => table.CellType(row, col).ToString(), row, col, null, string.Empty);
          cellDto.CellStyle = SafeGet(warnings, "cell.cellStyle", () => table.GetCellStyle(row, col), row, col, null, string.Empty);
          cellDto.CellState = SafeGet(warnings, "cell.cellState", () => table.GetCellState(row, col).ToString(), row, col, null, string.Empty);
          cellDto.CellStateValue = SafeGet(warnings, "cell.cellStateValue", () => (int)table.GetCellState(row, col), row, col, null, 0);
          cellDto.ToolTip = SafeGet(warnings, "cell.toolTip", () => table.GetToolTip(row, col), row, col, null, string.Empty);
          cellDto.Alignment = SafeGet(warnings, "cell.alignment", () => table.Alignment(row, col).ToString(), row, col, null, string.Empty);
          cellDto.TextRotation = SafeGet(warnings, "cell.textRotation", () => table.TextRotation(row, col).ToString(), row, col, null, string.Empty);
          cellDto.ContentLayout = SafeGet(warnings, "cell.contentLayout", () => table.GetContentLayout(row, col).ToString(), row, col, null, string.Empty);
          cellDto.DataType = SafeGet(warnings, "cell.dataType", () => table.DataType(row, col).ToString(), row, col, null, string.Empty);
          cellDto.UnitType = SafeGet(warnings, "cell.unitType", () => table.UnitType(row, col).ToString(), row, col, null, string.Empty);
          cellDto.DataFormat = SafeGet(warnings, "cell.dataFormat", () => table.GetDataFormat(row, col), row, col, null, string.Empty);
          cellDto.TopMargin = SafeGet<double?>(warnings, "cell.marginTop", () => table.GetMargin(row, col, CellMargins.Top), row, col, null, null);
          cellDto.LeftMargin = SafeGet<double?>(warnings, "cell.marginLeft", () => table.GetMargin(row, col, CellMargins.Left), row, col, null, null);
          cellDto.BottomMargin = SafeGet<double?>(warnings, "cell.marginBottom", () => table.GetMargin(row, col, CellMargins.Bottom), row, col, null, null);
          cellDto.RightMargin = SafeGet<double?>(warnings, "cell.marginRight", () => table.GetMargin(row, col, CellMargins.Right), row, col, null, null);
          cellDto.TextHeight = SafeGet<double?>(warnings, "cell.textHeight", () => table.TextHeight(row, col), row, col, null, null);
          cellDto.TextStyle = ToObjectRef(SafeGet(warnings, "cell.textStyleId", () => table.TextStyle(row, col), row, col, null, ObjectId.Null), tr, warnings, "cell.textStyle", row, col, null);
          cellDto.ContentColor = ToColorDto(SafeGet(warnings, "cell.contentColor", () => table.ContentColor(row, col), row, col, null, null as Color));
          cellDto.BackgroundColor = ToColorDto(SafeGet(warnings, "cell.backgroundColor", () => table.BackgroundColor(row, col), row, col, null, null as Color));
          cellDto.IsBackgroundColorNone = SafeGet<bool?>(warnings, "cell.isBackgroundColorNone", () => table.IsBackgroundColorNone(row, col), row, col, null, null);
          cellDto.Field = ToObjectRef(SafeGet(warnings, "cell.fieldId", () => table.FieldId(row, col), row, col, null, ObjectId.Null), tr, warnings, "cell.field", row, col, null);
          cellDto.DataLink = ToObjectRef(SafeGet(warnings, "cell.dataLink", () => table.GetDataLink(row, col), row, col, null, ObjectId.Null), tr, warnings, "cell.dataLinkObject", row, col, null);
          cellDto.IsLinked = cell?.IsLinked;
          cellDto.IsContentEditable = cell?.IsContentEditable;
          cellDto.IsFormatEditable = cell?.IsFormatEditable;
          cellDto.IsEmpty = cell?.IsEmpty;
          cellDto.StyleOverrides = ToEnumNameList(SafeGet(warnings, "cell.styleOverrides", () => table.CellStyleOverrides(row, col), row, col, null, new TableStyleOverride[0]));
          cellDto.StyleOverrideValues = ToEnumValueList(SafeGet(warnings, "cell.styleOverrideValues", () => table.CellStyleOverrides(row, col), row, col, null, new TableStyleOverride[0]));
          cellDto.OuterExtents = SafeGet(warnings, "cell.outerExtents", () => GetCellExtents(table, row, col, true), row, col, null, new List<TableAnalyzePoint>());
          cellDto.InnerExtents = SafeGet(warnings, "cell.innerExtents", () => GetCellExtents(table, row, col, false), row, col, null, new List<TableAnalyzePoint>());
          cellDto.GridLines = CaptureGridLines(table, tr, warnings, row, col);

          CaptureCellContents(table, tr, warnings, row, col, cell, cellDto);
          cells.Add(cellDto);
        }
      }
    }

    private static void CaptureCellMerge(
      Table table,
      List<TableAnalyzeWarning> warnings,
      int row,
      int col,
      TableAnalyzeCell cellDto,
      List<TableAnalyzeMergedRange> mergedRanges,
      HashSet<string> mergedRangeKeys
    )
    {
      CellRange mergeRange = null;
      bool isMerged = false;
      try
      {
        isMerged = table.IsMergedCell(row, col, out mergeRange);
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, "cell.isMergedCell", ex, row, col, null);
      }

      cellDto.IsMerged = isMerged;
      if (!isMerged || mergeRange == null)
      {
        return;
      }

      cellDto.MergeRange = ToCellRangeDto(mergeRange);
      string key = BuildMergeRangeKey(mergeRange);
      if (mergedRangeKeys.Add(key))
      {
        mergedRanges.Add(new TableAnalyzeMergedRange
        {
          TopRow = mergeRange.TopRow,
          LeftColumn = mergeRange.LeftColumn,
          BottomRow = mergeRange.BottomRow,
          RightColumn = mergeRange.RightColumn
        });
      }
    }

    private static void CaptureCellContents(
      Table table,
      Transaction tr,
      List<TableAnalyzeWarning> warnings,
      int row,
      int col,
      Cell cell,
      TableAnalyzeCell cellDto
    )
    {
      int contentCount = SafeGet(warnings, "cell.contentCount", () => table.GetNumberOfContents(row, col), row, col, null, 0);
      cellDto.ContentCount = contentCount;
      cellDto.Contents = new List<TableAnalyzeCellContent>(Math.Max(contentCount, 0));

      for (int contentIndex = 0; contentIndex < contentCount; contentIndex++)
      {
        var dto = new TableAnalyzeCellContent { ContentIndex = contentIndex };
        CellContent content = null;
        if (cell != null)
        {
          content = SafeGet(warnings, "content.object", () => cell.Contents[contentIndex], row, col, contentIndex, null as CellContent);
        }

        CellContentTypes contentTypes = SafeGet(warnings, "content.types", () => table.GetContentTypes(row, col, contentIndex), row, col, contentIndex, CellContentTypes.Unknown);
        dto.ContentTypes = contentTypes.ToString();
        dto.ContentTypesValue = (int)contentTypes;
        dto.TextString = SafeGet(warnings, "content.textString", () => table.GetTextString(row, col, contentIndex), row, col, contentIndex, string.Empty);
        dto.TextStringForEditing = SafeGet(warnings, "content.textForEditing", () => table.GetTextString(row, col, contentIndex, FormatOption.ForEditing), row, col, contentIndex, string.Empty);
        dto.TextStringForExpression = SafeGet(warnings, "content.textForExpression", () => table.GetTextString(row, col, contentIndex, FormatOption.ForExpression), row, col, contentIndex, string.Empty);
        dto.TextStringWithoutMtext = SafeGet(warnings, "content.textNoMText", () => table.GetTextString(row, col, contentIndex, FormatOption.IgnoreMtextFormat), row, col, contentIndex, string.Empty);

        dto.Value = ConvertValue(SafeGet<object>(warnings, "content.value", () => table.GetValue(row, col, contentIndex), row, col, contentIndex, null));
        dto.FormattedValue = ConvertValue(SafeGet<object>(warnings, "content.formattedValue", () => table.GetValue(row, col, contentIndex, FormatOption.ForEditing), row, col, contentIndex, null));
        dto.Formula = SafeGet(warnings, "content.formula", () => table.GetFormula(row, col, contentIndex), row, col, contentIndex, string.Empty);
        dto.HasFormula = SafeGet<bool?>(warnings, "content.hasFormula", () => table.HasFormula(row, col, contentIndex), row, col, contentIndex, null);
        dto.IsAutoScale = SafeGet<bool?>(warnings, "content.autoScale", () => table.GetIsAutoScale(row, col, contentIndex), row, col, contentIndex, null);
        dto.Rotation = SafeGet<double?>(warnings, "content.rotation", () => table.GetRotation(row, col, contentIndex), row, col, contentIndex, null);
        dto.Scale = SafeGet<double?>(warnings, "content.scale", () => table.GetScale(row, col, contentIndex), row, col, contentIndex, null);
        dto.TextHeight = SafeGet<double?>(warnings, "content.textHeight", () => table.GetTextHeight(row, col, contentIndex), row, col, contentIndex, null);
        dto.TextStyle = ToObjectRef(SafeGet(warnings, "content.textStyleId", () => table.GetTextStyleId(row, col, contentIndex), row, col, contentIndex, ObjectId.Null), tr, warnings, "content.textStyle", row, col, contentIndex);
        dto.ContentColor = ToColorDto(SafeGet(warnings, "content.color", () => table.GetContentColor(row, col, contentIndex), row, col, contentIndex, null as Color));

        DataTypeParameter dt = SafeGet(warnings, "content.dataType", () => table.GetDataType(row, col, contentIndex), row, col, contentIndex, default(DataTypeParameter));
        dto.DataType = dt.DataType.ToString();
        dto.UnitType = dt.UnitType.ToString();
        dto.DataFormat = SafeGet(warnings, "content.dataFormat", () => table.GetDataFormat(row, col, contentIndex), row, col, contentIndex, string.Empty);
        dto.Field = ToObjectRef(SafeGet(warnings, "content.fieldId", () => table.GetFieldId(row, col, contentIndex), row, col, contentIndex, ObjectId.Null), tr, warnings, "content.field", row, col, contentIndex);

        ObjectId btrId = SafeGet(warnings, "content.blockTableRecordId", () => table.GetBlockTableRecordId(row, col, contentIndex), row, col, contentIndex, ObjectId.Null);
        dto.BlockTableRecord = ToObjectRef(btrId, tr, warnings, "content.blockTableRecord", row, col, contentIndex);
        CellProperties overrides = SafeGet(warnings, "content.overrides", () => table.GetOverrides(row, col, contentIndex), row, col, contentIndex, CellProperties.Invalid);
        dto.OverrideFlags = overrides.ToString();
        dto.OverrideFlagValue = (int)overrides;
        dto.OverrideFlagNames = SplitEnumString(overrides.ToString());

        if (content != null)
        {
          dto.TextStringFromContentObject = SafeGet(warnings, "content.objectText", () => content.TextString, row, col, contentIndex, string.Empty);
          dto.ValueFromContentObject = ConvertValue(SafeGet<object>(warnings, "content.objectValue", () => content.Value, row, col, contentIndex, null));
          dto.HasFormulaFromContentObject = SafeGet<bool?>(warnings, "content.objectHasFormula", () => content.HasFormula, row, col, contentIndex, null);
          dto.FormulaFromContentObject = SafeGet(warnings, "content.objectFormula", () => content.Formula, row, col, contentIndex, string.Empty);
        }

        dto.BlockAttributeValues = (!btrId.IsNull && btrId.IsValid)
          ? CaptureBlockAttributeValues(table, tr, warnings, row, col, contentIndex, btrId)
          : new List<TableAnalyzeBlockAttributeValue>();

        cellDto.Contents.Add(dto);
      }
    }

    private static List<TableAnalyzeGridLine> CaptureGridLines(Table table, Transaction tr, List<TableAnalyzeWarning> warnings, int row, int col)
    {
      var gridLines = new List<TableAnalyzeGridLine>();
      foreach (GridLineType gridType in Enum.GetValues(typeof(GridLineType)))
      {
        if (gridType == GridLineType.InvalidGridLine) continue;

        Visibility visibility = SafeGet(warnings, "grid.visibility", () => table.GetGridVisibility(row, col, gridType), row, col, null, Visibility.Visible);
        GridLineStyle lineStyle = SafeGet(warnings, "grid.lineStyle", () => table.GetGridLineStyle(row, col, gridType), row, col, null, GridLineStyle.Single);
        LineWeight lineWeight = SafeGet(warnings, "grid.lineWeight", () => table.GetGridLineWeight(row, col, gridType), row, col, null, LineWeight.ByLineWeightDefault);
        ObjectId linetype = SafeGet(warnings, "grid.linetype", () => table.GetGridLinetype(row, col, gridType), row, col, null, ObjectId.Null);
        Color color = SafeGet(warnings, "grid.color", () => table.GetGridColor(row, col, gridType), row, col, null, null as Color);
        double? spacing = SafeGet<double?>(warnings, "grid.doubleLineSpacing", () => table.GetGridDoubleLineSpacing(row, col, gridType), row, col, null, null);
        GridPropertyParameter mask = SafeGet(warnings, "grid.propertyMask", () => table.GetGridProperty(row, col, gridType), row, col, null, default(GridPropertyParameter));
        CellProperties ov = SafeGet(warnings, "grid.overrides", () => table.GetOverrides(row, col, gridType), row, col, null, CellProperties.Invalid);

        gridLines.Add(new TableAnalyzeGridLine
        {
          GridLineType = gridType.ToString(),
          Visibility = visibility.ToString(),
          IsVisible = visibility == Visibility.Visible,
          LineStyle = lineStyle.ToString(),
          LineWeight = lineWeight.ToString(),
          LineWeightValue = (int)lineWeight,
          Linetype = ToObjectRef(linetype, tr, warnings, "grid.linetypeObject", row, col, null),
          Color = ToColorDto(color),
          DoubleLineSpacing = spacing,
          PropertyMask = mask.PropertyMask.ToString(),
          PropertyMaskValue = (int)mask.PropertyMask,
          PropertyMaskNames = SplitEnumString(mask.PropertyMask.ToString()),
          OverrideFlags = ov.ToString(),
          OverrideFlagValue = (int)ov,
          OverrideFlagNames = SplitEnumString(ov.ToString())
        });
      }
      return gridLines;
    }

    private static List<TableAnalyzeBlockAttributeValue> CaptureBlockAttributeValues(
      Table table,
      Transaction tr,
      List<TableAnalyzeWarning> warnings,
      int row,
      int col,
      int contentIndex,
      ObjectId blockTableRecordId
    )
    {
      var values = new List<TableAnalyzeBlockAttributeValue>();
      try
      {
        BlockTableRecord blockDef = tr.GetObject(blockTableRecordId, OpenMode.ForRead, false) as BlockTableRecord;
        if (blockDef == null) return values;

        foreach (ObjectId id in blockDef)
        {
          AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead, false) as AttributeDefinition;
          if (attDef == null) continue;
          string val = SafeGet(warnings, "content.blockAttribute", () => table.GetBlockAttributeValue(row, col, contentIndex, id), row, col, contentIndex, string.Empty);
          values.Add(new TableAnalyzeBlockAttributeValue
          {
            Tag = attDef.Tag,
            Prompt = attDef.Prompt,
            Value = val,
            AttributeDefinition = ToObjectRef(id, tr, warnings, "content.attributeDef", row, col, contentIndex)
          });
        }
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, "content.blockAttributeCapture", ex, row, col, contentIndex);
      }
      return values;
    }

    private static string ResolveSpaceName(Table table, Transaction tr, List<TableAnalyzeWarning> warnings)
    {
      try
      {
        BlockTableRecord owner = tr.GetObject(table.OwnerId, OpenMode.ForRead, false) as BlockTableRecord;
        if (owner == null) return "Unknown";
        if (string.Equals(owner.Name, BlockTableRecord.ModelSpace, StringComparison.OrdinalIgnoreCase)) return "ModelSpace";
        return owner.IsLayout ? "PaperSpace" : owner.Name ?? "Unknown";
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, "source.space", ex, null, null, null);
        return "Unknown";
      }
    }

    private static string WriteExportToJson(TableAnalyzeExport export, Database db)
    {
      string dwgPath = db.Filename ?? string.Empty;
      string outputDir = string.IsNullOrWhiteSpace(dwgPath)
        ? Path.Combine(Path.GetTempPath(), "ElectricalCommands", "TableAnalysis")
        : Path.GetDirectoryName(dwgPath);
      if (string.IsNullOrWhiteSpace(outputDir))
      {
        outputDir = Path.Combine(Path.GetTempPath(), "ElectricalCommands", "TableAnalysis");
      }

      Directory.CreateDirectory(outputDir);
      string safeHandle = SanitizeFileNamePart(string.IsNullOrWhiteSpace(export.Source?.TableHandle) ? "UNKNOWN" : export.Source.TableHandle);
      string fileName = $"TABLEANALYZE_{safeHandle}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
      string path = Path.Combine(outputDir, fileName);

      JsonSerializerSettings settings = new JsonSerializerSettings
      {
        Formatting = Formatting.Indented,
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        NullValueHandling = NullValueHandling.Ignore
      };
      File.WriteAllText(path, JsonConvert.SerializeObject(export, settings));
      return path;
    }

    private static string SanitizeFileNamePart(string value)
    {
      if (string.IsNullOrWhiteSpace(value)) return "UNKNOWN";
      char[] invalid = Path.GetInvalidFileNameChars();
      return new string(value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
    }

    private static List<TableAnalyzePoint> GetCellExtents(Table table, int row, int col, bool isOuterCell)
    {
      Point3dCollection pts = new Point3dCollection();
      table.GetCellExtents(row, col, isOuterCell, pts);
      return pts.Cast<Point3d>().Select(ToPointDto).ToList();
    }

    private static string BuildMergeRangeKey(CellRange range)
    {
      return $"{range.TopRow}:{range.LeftColumn}:{range.BottomRow}:{range.RightColumn}";
    }

    private static TableAnalyzeCellRange ToCellRangeDto(CellRange range)
    {
      if (range == null) return null;
      return new TableAnalyzeCellRange
      {
        TopRow = range.TopRow,
        LeftColumn = range.LeftColumn,
        BottomRow = range.BottomRow,
        RightColumn = range.RightColumn
      };
    }

    private static TableAnalyzeObjectRef ToObjectRef(
      ObjectId id,
      Transaction tr,
      List<TableAnalyzeWarning> warnings,
      string property,
      int? row = null,
      int? col = null,
      int? contentIndex = null
    )
    {
      var dto = new TableAnalyzeObjectRef { ObjectId = id.ToString(), IsNull = id.IsNull };
      if (id.IsNull || !id.IsValid) return dto;
      try
      {
        DBObject obj = tr.GetObject(id, OpenMode.ForRead, false);
        if (obj == null) return dto;
        dto.Handle = obj.Handle.ToString();
        dto.ClassName = obj.GetRXClass()?.Name ?? obj.GetType().Name;
        dto.Name = TryGetObjectName(obj);
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, property, ex, row, col, contentIndex);
      }
      return dto;
    }

    private static string TryGetObjectName(DBObject obj)
    {
      if (obj is SymbolTableRecord sym) return sym.Name;
      var prop = obj.GetType().GetProperty("Name");
      if (prop != null && prop.PropertyType == typeof(string)) return prop.GetValue(obj, null) as string;
      return null;
    }

    private static TableAnalyzeColor ToColorDto(Color color)
    {
      if (color == null) return null;
      return new TableAnalyzeColor
      {
        ColorMethod = color.ColorMethod.ToString(),
        ColorIndex = color.ColorIndex,
        Red = color.Red,
        Green = color.Green,
        Blue = color.Blue,
        ColorName = color.ColorName,
        BookName = color.BookName,
        IsByAci = color.IsByAci,
        IsByLayer = color.IsByLayer,
        IsByBlock = color.IsByBlock,
        IsForeground = color.IsForeground,
        IsNone = color.IsNone
      };
    }

    private static TableAnalyzePoint ToPointDto(Point3d point)
    {
      return new TableAnalyzePoint { X = point.X, Y = point.Y, Z = point.Z };
    }

    private static TableAnalyzeVector ToVectorDto(Vector3d vector)
    {
      return new TableAnalyzeVector { X = vector.X, Y = vector.Y, Z = vector.Z };
    }

    private static TableAnalyzeValue ConvertValue(object value)
    {
      if (value == null)
      {
        return new TableAnalyzeValue { Kind = "null", TypeName = "null", PrimitiveValue = null, DisplayValue = null };
      }
      Type type = value.GetType();
      string typeName = type.FullName ?? type.Name;
      if (value is string || value is bool || value is byte || value is sbyte || value is short || value is ushort || value is int || value is uint || value is long || value is ulong || value is float || value is double || value is decimal)
      {
        return new TableAnalyzeValue
        {
          Kind = "primitive",
          TypeName = typeName,
          PrimitiveValue = value,
          DisplayValue = Convert.ToString(value, CultureInfo.InvariantCulture)
        };
      }
      if (value is DateTime dt)
      {
        string iso = dt.ToString("o", CultureInfo.InvariantCulture);
        return new TableAnalyzeValue { Kind = "dateTime", TypeName = typeName, PrimitiveValue = iso, DisplayValue = iso };
      }
      if (value is IFormattable fmt)
      {
        string formatted = fmt.ToString(null, CultureInfo.InvariantCulture);
        return new TableAnalyzeValue { Kind = "formattable", TypeName = typeName, PrimitiveValue = formatted, DisplayValue = formatted };
      }
      string fallback = value.ToString();
      return new TableAnalyzeValue { Kind = "stringified", TypeName = typeName, PrimitiveValue = fallback, DisplayValue = fallback };
    }

    private static List<string> ToEnumNameList(IEnumerable<TableStyleOverride> values)
    {
      return values?.Select(v => v.ToString()).Distinct(StringComparer.OrdinalIgnoreCase).ToList() ?? new List<string>();
    }

    private static List<int> ToEnumValueList(IEnumerable<TableStyleOverride> values)
    {
      return values?.Select(v => (int)v).Distinct().ToList() ?? new List<int>();
    }

    private static List<string> SplitEnumString(string value)
    {
      if (string.IsNullOrWhiteSpace(value)) return new List<string>();
      return value
        .Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries)
        .Select(part => part.Trim())
        .Where(part => !string.IsNullOrWhiteSpace(part))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();
    }

    private static T SafeGet<T>(
      List<TableAnalyzeWarning> warnings,
      string property,
      Func<T> getter,
      int? row,
      int? col,
      int? contentIndex,
      T fallback = default(T)
    )
    {
      try
      {
        return getter();
      }
      catch (System.Exception ex)
      {
        AddWarning(warnings, property, ex, row, col, contentIndex);
        return fallback;
      }
    }

    private static void AddWarning(
      List<TableAnalyzeWarning> warnings,
      string property,
      System.Exception ex,
      int? row,
      int? col,
      int? contentIndex
    )
    {
      warnings.Add(new TableAnalyzeWarning
      {
        Property = property,
        Message = ex.Message,
        ExceptionType = ex.GetType().Name,
        Row = row,
        Column = col,
        ContentIndex = contentIndex
      });
    }
  }

  public class TableAnalyzeExport
  {
    public string SchemaVersion { get; set; }
    public DateTime CapturedAtUtc { get; set; }
    public TableAnalyzeSource Source { get; set; }
    public TableAnalyzeTable Table { get; set; }
    public List<TableAnalyzeMergedRange> MergedRanges { get; set; } = new List<TableAnalyzeMergedRange>();
    public List<TableAnalyzeCell> Cells { get; set; } = new List<TableAnalyzeCell>();
    public List<TableAnalyzeWarning> Warnings { get; set; } = new List<TableAnalyzeWarning>();
  }

  public class TableAnalyzeSource
  {
    public string DwgPath { get; set; }
    public string DrawingName { get; set; }
    public string TableHandle { get; set; }
    public string TableObjectId { get; set; }
    public string Layer { get; set; }
    public string Space { get; set; }
    public string CaptureToolVersion { get; set; }
  }

  public class TableAnalyzeTable
  {
    public TableAnalyzeObjectRef TableStyle { get; set; }
    public string TableStyleName { get; set; }
    public List<string> TableStyleOverrides { get; set; } = new List<string>();
    public List<int> TableStyleOverrideValues { get; set; } = new List<int>();
    public int NumRows { get; set; }
    public int NumColumns { get; set; }
    public List<double> RowHeights { get; set; } = new List<double>();
    public List<string> RowTypes { get; set; } = new List<string>();
    public List<double> ColumnWidths { get; set; } = new List<double>();
    public List<string> ColumnNames { get; set; } = new List<string>();
    public TableAnalyzePoint Position { get; set; }
    public TableAnalyzeVector Direction { get; set; }
    public TableAnalyzeVector Normal { get; set; }
    public double Rotation { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public double MinimumTableWidth { get; set; }
    public double MinimumTableHeight { get; set; }
    public double HorizontalCellMargin { get; set; }
    public double VerticalCellMargin { get; set; }
    public string FlowDirection { get; set; }
    public string MergeStyle { get; set; }
    public bool IsTitleSuppressed { get; set; }
    public bool IsHeaderSuppressed { get; set; }
    public bool BreakEnabled { get; set; }
  }

  public class TableAnalyzeMergedRange
  {
    public int TopRow { get; set; }
    public int LeftColumn { get; set; }
    public int BottomRow { get; set; }
    public int RightColumn { get; set; }
  }

  public class TableAnalyzeCell
  {
    public int Row { get; set; }
    public int Column { get; set; }
    public bool IsMerged { get; set; }
    public TableAnalyzeCellRange MergeRange { get; set; }
    public string RowType { get; set; }
    public string CellType { get; set; }
    public string CellStyle { get; set; }
    public string CellState { get; set; }
    public int CellStateValue { get; set; }
    public string ToolTip { get; set; }
    public string Alignment { get; set; }
    public string TextRotation { get; set; }
    public string ContentLayout { get; set; }
    public string UnitType { get; set; }
    public string DataType { get; set; }
    public string DataFormat { get; set; }
    public double? TopMargin { get; set; }
    public double? LeftMargin { get; set; }
    public double? BottomMargin { get; set; }
    public double? RightMargin { get; set; }
    public double? TextHeight { get; set; }
    public TableAnalyzeObjectRef TextStyle { get; set; }
    public TableAnalyzeColor ContentColor { get; set; }
    public TableAnalyzeColor BackgroundColor { get; set; }
    public bool? IsBackgroundColorNone { get; set; }
    public TableAnalyzeObjectRef Field { get; set; }
    public TableAnalyzeObjectRef DataLink { get; set; }
    public bool? IsLinked { get; set; }
    public bool? IsContentEditable { get; set; }
    public bool? IsFormatEditable { get; set; }
    public bool? IsEmpty { get; set; }
    public List<string> StyleOverrides { get; set; } = new List<string>();
    public List<int> StyleOverrideValues { get; set; } = new List<int>();
    public List<TableAnalyzePoint> OuterExtents { get; set; } = new List<TableAnalyzePoint>();
    public List<TableAnalyzePoint> InnerExtents { get; set; } = new List<TableAnalyzePoint>();
    public List<TableAnalyzeGridLine> GridLines { get; set; } = new List<TableAnalyzeGridLine>();
    public int ContentCount { get; set; }
    public List<TableAnalyzeCellContent> Contents { get; set; } = new List<TableAnalyzeCellContent>();
  }

  public class TableAnalyzeCellRange
  {
    public int TopRow { get; set; }
    public int LeftColumn { get; set; }
    public int BottomRow { get; set; }
    public int RightColumn { get; set; }
  }

  public class TableAnalyzeGridLine
  {
    public string GridLineType { get; set; }
    public string Visibility { get; set; }
    public bool IsVisible { get; set; }
    public string LineStyle { get; set; }
    public string LineWeight { get; set; }
    public int LineWeightValue { get; set; }
    public TableAnalyzeObjectRef Linetype { get; set; }
    public TableAnalyzeColor Color { get; set; }
    public double? DoubleLineSpacing { get; set; }
    public string PropertyMask { get; set; }
    public int PropertyMaskValue { get; set; }
    public List<string> PropertyMaskNames { get; set; } = new List<string>();
    public string OverrideFlags { get; set; }
    public int OverrideFlagValue { get; set; }
    public List<string> OverrideFlagNames { get; set; } = new List<string>();
  }

  public class TableAnalyzeCellContent
  {
    public int ContentIndex { get; set; }
    public string ContentTypes { get; set; }
    public int ContentTypesValue { get; set; }
    public string TextString { get; set; }
    public string TextStringForEditing { get; set; }
    public string TextStringForExpression { get; set; }
    public string TextStringWithoutMtext { get; set; }
    public string TextStringFromContentObject { get; set; }
    public TableAnalyzeValue Value { get; set; }
    public TableAnalyzeValue FormattedValue { get; set; }
    public TableAnalyzeValue ValueFromContentObject { get; set; }
    public string Formula { get; set; }
    public bool? HasFormula { get; set; }
    public string FormulaFromContentObject { get; set; }
    public bool? HasFormulaFromContentObject { get; set; }
    public bool? IsAutoScale { get; set; }
    public double? Rotation { get; set; }
    public double? Scale { get; set; }
    public double? TextHeight { get; set; }
    public TableAnalyzeObjectRef TextStyle { get; set; }
    public TableAnalyzeColor ContentColor { get; set; }
    public string DataType { get; set; }
    public string UnitType { get; set; }
    public string DataFormat { get; set; }
    public TableAnalyzeObjectRef Field { get; set; }
    public TableAnalyzeObjectRef BlockTableRecord { get; set; }
    public string OverrideFlags { get; set; }
    public int OverrideFlagValue { get; set; }
    public List<string> OverrideFlagNames { get; set; } = new List<string>();
    public List<TableAnalyzeBlockAttributeValue> BlockAttributeValues { get; set; } = new List<TableAnalyzeBlockAttributeValue>();
  }

  public class TableAnalyzeBlockAttributeValue
  {
    public string Tag { get; set; }
    public string Prompt { get; set; }
    public string Value { get; set; }
    public TableAnalyzeObjectRef AttributeDefinition { get; set; }
  }

  public class TableAnalyzeObjectRef
  {
    public string ObjectId { get; set; }
    public string Handle { get; set; }
    public string Name { get; set; }
    public string ClassName { get; set; }
    public bool IsNull { get; set; }
  }

  public class TableAnalyzeColor
  {
    public string ColorMethod { get; set; }
    public short ColorIndex { get; set; }
    public byte Red { get; set; }
    public byte Green { get; set; }
    public byte Blue { get; set; }
    public string ColorName { get; set; }
    public string BookName { get; set; }
    public bool IsByAci { get; set; }
    public bool IsByLayer { get; set; }
    public bool IsByBlock { get; set; }
    public bool IsForeground { get; set; }
    public bool IsNone { get; set; }
  }

  public class TableAnalyzePoint
  {
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }
  }

  public class TableAnalyzeVector
  {
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }
  }

  public class TableAnalyzeValue
  {
    public string Kind { get; set; }
    public string TypeName { get; set; }
    public object PrimitiveValue { get; set; }
    public string DisplayValue { get; set; }
  }

  public class TableAnalyzeWarning
  {
    public int? Row { get; set; }
    public int? Column { get; set; }
    public int? ContentIndex { get; set; }
    public string Property { get; set; }
    public string Message { get; set; }
    public string ExceptionType { get; set; }
  }
}
