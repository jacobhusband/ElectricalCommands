using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Globalization;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string HomerunTextStyleName = "ARIALNARROW-1-8";
    private const string HomerunTextStyleFont = "ARIALN.TTF";
    private const double QuarterScaleSymbolSize = 4.5;
    private const double QuarterScalePaperInchesPerFoot = 0.25;

    [CommandMethod("HRSETTINGS", CommandFlags.Modal)]
    [CommandMethod("HRS", CommandFlags.Modal)]
    public void ShowHomerunSettingsCommand()
    {
      HomerunSettingsPalette.Show();
    }

    [CommandMethod("SETPANELLOCATION", CommandFlags.Modal)]
    [CommandMethod("SPL", CommandFlags.Modal)]
    public void SetPanelLocationCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      PromptPointResult result = ed.GetPoint(
        "\nSelect the panel location in modelspace or the current paperspace viewport: "
      );
      if (result.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nSETPANELLOCATION canceled.");
        return;
      }

      try
      {
        string context = GetCurrentDrawingContext(db);
        string spaceHandle = db.CurrentSpaceId.Handle.ToString();
        ElectricalDrawingSettingsStore.WritePanelLocation(
          db,
          result.Value,
          spaceHandle,
          context
        );
        ed.WriteMessage($"\nPanel location set to {FormatPoint(result.Value)} ({context}).");
        HomerunSettingsPalette.Refresh();
        HomerunSettingsPalette.SetStatus("Panel location updated.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to save the panel location: {ex.Message}");
      }
    }

    [CommandMethod("SETSCALE", CommandFlags.Modal)]
    [CommandMethod("SS", CommandFlags.Modal)]
    public void SetScaleCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      string current = string.Empty;
      if (ElectricalDrawingSettingsStore.TryReadScale(db, out var existing))
      {
        current = $" Current scale: {existing.DisplayText}.";
      }

      PromptStringOptions options = new PromptStringOptions(
        $"\nEnter drawing scale (for example 1/4, 1/4\"=1'-0\", or 1:48).{current} "
      )
      {
        AllowSpaces = true,
      };
      PromptResult result = ed.GetString(options);
      if (result.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nSETSCALE canceled; the previous scale was not changed.");
        return;
      }

      if (!TryParseDrawingScale(
        result.StringResult,
        out double paperInchesPerFoot,
        out string displayText
      ))
      {
        ed.WriteMessage(
          "\nInvalid scale. Use a value such as 1/4, 3/16, 1/4\"=1'-0\", or 1:48."
        );
        return;
      }

      try
      {
        ElectricalDrawingSettingsStore.WriteScale(db, paperInchesPerFoot, displayText);
        double symbolSize = ResolveHomerunSymbolSize(paperInchesPerFoot);
        ed.WriteMessage(
          $"\nDrawing scale set to {displayText}. New HR objects will use a {FormatNumber(symbolSize)}\" arrow and text height; existing objects are unchanged."
        );
        HomerunSettingsPalette.Refresh();
        HomerunSettingsPalette.SetStatus($"Scale set to {displayText}.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to save the drawing scale: {ex.Message}");
      }
    }

    [CommandMethod("SETPANELNAME", CommandFlags.Modal)]
    [CommandMethod("SPN", CommandFlags.Modal)]
    public void SetPanelNameCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      string current = string.Empty;
      if (ElectricalDrawingSettingsStore.TryReadPanelName(db, out string existing))
      {
        current = $" Current panel: {existing}.";
      }

      PromptStringOptions options = new PromptStringOptions(
        $"\nEnter the panel name other commands should use.{current} "
      )
      {
        AllowSpaces = true,
      };
      PromptResult result = ed.GetString(options);
      if (result.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nSETPANELNAME canceled; the previous panel name was not changed.");
        return;
      }

      string panelName = (result.StringResult ?? string.Empty).Trim();
      if (panelName.Length == 0)
      {
        ed.WriteMessage("\nPanel name cannot be blank.");
        return;
      }

      try
      {
        ElectricalDrawingSettingsStore.WritePanelName(db, panelName);
        ed.WriteMessage($"\nPanel name set to {panelName}. HR labels will start with {BuildPanelLabel(panelName)}");
        HomerunSettingsPalette.Refresh();
        HomerunSettingsPalette.SetStatus($"Panel name set to {panelName}.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to save the panel name: {ex.Message}");
      }
    }

    [CommandMethod("HOMERUN", CommandFlags.Modal)]
    [CommandMethod("HR", CommandFlags.Modal)]
    public void HomerunCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!ElectricalDrawingSettingsStore.TryReadPanelLocation(db, out var panelLocation))
      {
        ed.WriteMessage("\nHR requires a panel location. Run SETPANELLOCATION (SPL) first.");
        return;
      }
      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage("\nHR requires a drawing scale. Run SETSCALE (SS) first.");
        return;
      }
      if (!ElectricalDrawingSettingsStore.TryReadPanelName(db, out string panelName))
      {
        ed.WriteMessage("\nHR requires a panel name. Run SETPANELNAME (SPN) first.");
        return;
      }

      double symbolSize = ResolveHomerunSymbolSize(scale.PaperInchesPerModelFoot);
      string panelLabel = BuildPanelLabel(panelName);
      ObjectId targetLayerId = ResolveHomerunLayerId(db, out string targetLayerName);
      string currentSpaceHandle = db.CurrentSpaceId.Handle.ToString();
      if (!string.Equals(
        panelLocation.SpaceHandle,
        currentSpaceHandle,
        StringComparison.OrdinalIgnoreCase
      ))
      {
        ed.WriteMessage(
          $"\nHR requires the panel location in the current drawing space so the arrow preview can point toward it. The saved location is in {panelLocation.Context}; run SPL here and try again."
        );
        return;
      }

      Point3d tipPoint;
      Point3d middlePoint;
      Point3d basePoint;
      Point3d textPoint;
      try
      {
        using (HomerunArrowTipJig tipJig = new HomerunArrowTipJig(
          db,
          panelLocation.Point,
          symbolSize
        ))
        {
          PromptResult tipResult = ed.Drag(tipJig);
          if (tipResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nHR canceled.");
            return;
          }
          tipPoint = tipJig.TipPoint;
        }

        using (HomerunMiddleJig middleJig = new HomerunMiddleJig(
          db,
          tipPoint,
          panelLocation.Point,
          symbolSize
        ))
        {
          PromptResult middleResult = ed.Drag(middleJig);
          if (middleResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nHR canceled.");
            return;
          }
          middlePoint = middleJig.MiddlePoint;
        }

        using (HomerunBaseJig baseJig = new HomerunBaseJig(
          db,
          tipPoint,
          middlePoint,
          symbolSize
        ))
        {
          PromptResult baseResult = ed.Drag(baseJig);
          if (baseResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nHR canceled.");
            return;
          }
          basePoint = baseJig.BasePoint;
        }

        ObjectId previewTextStyleId = ResolveExistingTextStyleId(db, HomerunTextStyleName);
        using (HomerunTextJig textJig = new HomerunTextJig(
          db,
          tipPoint,
          middlePoint,
          basePoint,
          basePoint,
          symbolSize,
          panelLabel,
          previewTextStyleId
        ))
        {
          PromptResult textResult = ed.Drag(textJig);
          if (textResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nHR canceled.");
            return;
          }
          textPoint = textJig.TextPoint;
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to preview the home run: {ex.Message}");
        return;
      }

      try
      {
        ObjectId textStyleId = EnsureHomerunTextStyle(db);
        using (Transaction transaction = db.TransactionManager.StartTransaction())
        {
          BlockTableRecord space = transaction.GetObject(
            db.CurrentSpaceId,
            OpenMode.ForWrite
          ) as BlockTableRecord;
          if (space == null)
          {
            throw new InvalidOperationException("Unable to open the current drawing space.");
          }

          MText text = HomerunEntityFactory.CreateMText(
            db,
            textPoint,
            symbolSize,
            panelLabel,
            textStyleId
          );
          text.LayerId = targetLayerId;
          space.AppendEntity(text);
          transaction.AddNewlyCreatedDBObject(text, true);

          Leader leader = HomerunEntityFactory.CreateLeader(
            db,
            tipPoint,
            middlePoint,
            basePoint,
            symbolSize
          );
          leader.LayerId = targetLayerId;
          space.AppendEntity(leader);
          transaction.AddNewlyCreatedDBObject(leader, true);

          transaction.Commit();
        }

        ed.Regen();
        ed.WriteMessage(
          $"\nHR created for panel {panelName} at {scale.DisplayText} on layer {targetLayerName} with {FormatNumber(symbolSize)}\" text and arrow size."
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to create the home run: {ex.Message}");
      }
    }

    private static double ResolveHomerunSymbolSize(double paperInchesPerModelFoot)
    {
      return QuarterScaleSymbolSize
        * QuarterScalePaperInchesPerFoot
        / paperInchesPerModelFoot;
    }

    internal static ObjectId ResolveHomerunLayerId(Database database, out string layerName)
    {
      string storedLayer = string.Empty;
      ElectricalDrawingSettingsStore.TryReadHomerunLayer(database, out storedLayer);

      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        LayerTable layers = transaction.GetObject(
          database.LayerTableId,
          OpenMode.ForRead
        ) as LayerTable;
        if (layers == null)
        {
          layerName = string.Empty;
          return database.Clayer;
        }

        ObjectId layerId = !string.IsNullOrWhiteSpace(storedLayer) && layers.Has(storedLayer)
          ? layers[storedLayer]
          : database.Clayer;
        LayerTableRecord layer = transaction.GetObject(
          layerId,
          OpenMode.ForRead
        ) as LayerTableRecord;
        layerName = layer?.Name ?? storedLayer;
        return layerId;
      }
    }

    private static string BuildPanelLabel(string panelName)
    {
      string value = (panelName ?? string.Empty).Trim();
      return value.EndsWith("-", StringComparison.Ordinal) ? value : value + "-";
    }

    private static ObjectId ResolveExistingTextStyleId(Database database, string styleName)
    {
      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        TextStyleTable styles = transaction.GetObject(
          database.TextStyleTableId,
          OpenMode.ForRead
        ) as TextStyleTable;
        if (styles != null && styles.Has(styleName))
        {
          return styles[styleName];
        }
        return database.Textstyle;
      }
    }

    private static ObjectId EnsureHomerunTextStyle(Database database)
    {
      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        TextStyleTable styles = transaction.GetObject(
          database.TextStyleTableId,
          OpenMode.ForRead
        ) as TextStyleTable;
        if (styles == null)
        {
          throw new InvalidOperationException("Unable to open the text style table.");
        }
        if (styles.Has(HomerunTextStyleName))
        {
          ObjectId existingId = styles[HomerunTextStyleName];
          transaction.Commit();
          return existingId;
        }

        styles.UpgradeOpen();
        TextStyleTableRecord style = new TextStyleTableRecord
        {
          Name = HomerunTextStyleName,
          FileName = HomerunTextStyleFont,
          BigFontFileName = string.Empty,
          XScale = 1.0,
          ObliquingAngle = 0.0,
          TextSize = 0.0,
        };
        ObjectId id = styles.Add(style);
        transaction.AddNewlyCreatedDBObject(style, true);
        transaction.Commit();
        return id;
      }
    }

    private static string GetCurrentDrawingContext(Database database)
    {
      if (database.TileMode)
      {
        return "modelspace";
      }

      try
      {
        short viewportNumber = Convert.ToInt16(Application.GetSystemVariable("CVPORT"));
        if (viewportNumber > 1)
        {
          return $"modelspace through viewport on {LayoutManager.Current.CurrentLayout}";
        }
      }
      catch { }

      return $"paperspace layout {LayoutManager.Current.CurrentLayout}";
    }

    private static string FormatPoint(Point3d point)
    {
      return $"({FormatNumber(point.X)}, {FormatNumber(point.Y)}, {FormatNumber(point.Z)})";
    }

    private static string FormatNumber(double value)
    {
      return value.ToString("0.####", CultureInfo.InvariantCulture);
    }

    internal static bool TryParseDrawingScale(
      string input,
      out double paperInchesPerModelFoot,
      out string displayText
    )
    {
      paperInchesPerModelFoot = 0.0;
      displayText = string.Empty;
      string value = NormalizeScaleText(input);
      if (value.Length == 0)
      {
        return false;
      }

      int equalsIndex = value.IndexOf('=');
      if (equalsIndex >= 0)
      {
        if (value.IndexOf('=', equalsIndex + 1) >= 0)
        {
          return false;
        }

        string paperPart = value.Substring(0, equalsIndex);
        string modelPart = value.Substring(equalsIndex + 1);
        if (!TryParsePaperInches(paperPart, out double paperInches)
          || !TryParseArchitecturalLengthInches(modelPart, out double modelInches)
          || paperInches <= 0.0
          || modelInches <= 0.0)
        {
          return false;
        }

        paperInchesPerModelFoot = 12.0 * paperInches / modelInches;
      }
      else
      {
        int colonIndex = value.IndexOf(':');
        if (colonIndex >= 0)
        {
          if (value.IndexOf(':', colonIndex + 1) >= 0
            || !TryParseNumberOrFraction(value.Substring(0, colonIndex), out double paperUnits)
            || !TryParseNumberOrFraction(value.Substring(colonIndex + 1), out double modelUnits)
            || paperUnits <= 0.0
            || modelUnits <= 0.0)
          {
            return false;
          }
          paperInchesPerModelFoot = 12.0 * paperUnits / modelUnits;
        }
        else if (!TryParsePaperInches(value, out paperInchesPerModelFoot))
        {
          return false;
        }
      }

      if (paperInchesPerModelFoot <= 0.0
        || double.IsNaN(paperInchesPerModelFoot)
        || double.IsInfinity(paperInchesPerModelFoot))
      {
        return false;
      }

      displayText = FormatArchitecturalScale(paperInchesPerModelFoot);
      return true;
    }

    private static string NormalizeScaleText(string input)
    {
      return (input ?? string.Empty)
        .Trim()
        .Replace(" ", string.Empty)
        .Replace("\t", string.Empty)
        .Replace('’', '\'')
        .Replace('′', '\'')
        .Replace('“', '"')
        .Replace('”', '"')
        .Replace('″', '"');
    }

    private static bool TryParsePaperInches(string value, out double inches)
    {
      string cleaned = (value ?? string.Empty)
        .Replace("inches", string.Empty)
        .Replace("inch", string.Empty)
        .Replace("in", string.Empty)
        .Replace("\"", string.Empty);
      return TryParseNumberOrFraction(cleaned, out inches) && inches > 0.0;
    }

    private static bool TryParseArchitecturalLengthInches(string value, out double inches)
    {
      inches = 0.0;
      string cleaned = (value ?? string.Empty)
        .Replace("feet", "'")
        .Replace("foot", "'")
        .Replace("ft", "'")
        .Replace("inches", "\"")
        .Replace("inch", "\"")
        .Replace("in", "\"");

      int feetMark = cleaned.IndexOf('\'');
      if (feetMark >= 0)
      {
        string feetPart = cleaned.Substring(0, feetMark);
        string inchPart = cleaned.Substring(feetMark + 1).TrimStart('-').TrimEnd('"');
        if (!TryParseNumberOrFraction(feetPart, out double feet) || feet < 0.0)
        {
          return false;
        }
        double remainder = 0.0;
        if (inchPart.Length > 0
          && (!TryParseNumberOrFraction(inchPart, out remainder) || remainder < 0.0))
        {
          return false;
        }
        inches = feet * 12.0 + remainder;
        return inches > 0.0;
      }

      cleaned = cleaned.TrimEnd('"');
      return TryParseNumberOrFraction(cleaned, out inches) && inches > 0.0;
    }

    private static bool TryParseNumberOrFraction(string value, out double result)
    {
      result = 0.0;
      string cleaned = (value ?? string.Empty).Trim();
      int slash = cleaned.IndexOf('/');
      if (slash < 0)
      {
        return double.TryParse(
          cleaned,
          NumberStyles.Float,
          CultureInfo.InvariantCulture,
          out result
        );
      }
      if (cleaned.IndexOf('/', slash + 1) >= 0)
      {
        return false;
      }

      return double.TryParse(
          cleaned.Substring(0, slash),
          NumberStyles.Float,
          CultureInfo.InvariantCulture,
          out double numerator
        )
        && double.TryParse(
          cleaned.Substring(slash + 1),
          NumberStyles.Float,
          CultureInfo.InvariantCulture,
          out double denominator
        )
        && Math.Abs(denominator) > 1e-12
        && (result = numerator / denominator) >= 0.0;
    }

    private static string FormatArchitecturalScale(double paperInchesPerModelFoot)
    {
      string[] labels = { "1", "3/4", "1/2", "3/8", "1/4", "3/16", "1/8", "3/32", "1/16" };
      double[] values = { 1.0, 0.75, 0.5, 0.375, 0.25, 0.1875, 0.125, 0.09375, 0.0625 };
      for (int i = 0; i < values.Length; i++)
      {
        if (Math.Abs(paperInchesPerModelFoot - values[i]) < 1e-9)
        {
          return $"{labels[i]}\" = 1'-0\"";
        }
      }
      return $"{FormatNumber(paperInchesPerModelFoot)}\" = 1'-0\"";
    }
  }
}
