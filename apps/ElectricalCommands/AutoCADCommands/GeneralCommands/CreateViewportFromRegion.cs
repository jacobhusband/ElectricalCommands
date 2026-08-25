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
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const double ViewportAutoFitMaximumWidth = 30.0;
    private const double ViewportAutoFitMaximumHeight = 22.0;

    [CommandMethod("VPFROMREG")]
    [CommandMethod("QVP")]
    public void CREATEVIEWPORTFROMREGION()
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      PromptPointResult pointResult1 = ed.GetPoint(
        new PromptPointOptions("\nSelect the first corner of the region in modelspace: ")
      );
      if (pointResult1.Status != PromptStatus.OK)
      {
        return;
      }

      PromptPointOptions pointOptions2 = new PromptPointOptions(
        "\nSelect the opposite corner of the region in modelspace: "
      )
      {
        BasePoint = pointResult1.Value,
        UseBasePoint = true
      };
      PromptPointResult pointResult2 = ed.GetPoint(pointOptions2);
      if (pointResult2.Status != PromptStatus.OK)
      {
        return;
      }

      var correctedPoints = GetCorrectedPoints(pointResult1.Value, pointResult2.Value);
      Extents3d regionExtents = new Extents3d(correctedPoints.Min, correctedPoints.Max);
      double regionWidth = regionExtents.MaxPoint.X - regionExtents.MinPoint.X;
      double regionHeight = regionExtents.MaxPoint.Y - regionExtents.MinPoint.Y;

      if (regionWidth <= Tolerance.Global.EqualPoint || regionHeight <= Tolerance.Global.EqualPoint)
      {
        ed.WriteMessage("\nVPFROMREG: The selected region must have a non-zero width and height.");
        return;
      }

      List<string> paperLayouts = GetPaperLayoutNames(db);
      if (paperLayouts.Count == 0)
      {
        ed.WriteMessage("\nVPFROMREG: No paperspace layouts were found in the drawing.");
        return;
      }

      List<ViewportScaleOption> scaleOptions = CreateViewportScaleOptions();
      string currentLayoutName = LayoutManager.Current.CurrentLayout;
      ViewportFromRegionOptionsWindow optionsWindow = new ViewportFromRegionOptionsWindow(
        paperLayouts,
        scaleOptions,
        currentLayoutName,
        regionWidth,
        regionHeight,
        ViewportAutoFitMaximumWidth,
        ViewportAutoFitMaximumHeight
      );

      bool? accepted = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(
        optionsWindow
      );
      if (accepted != true)
      {
        ed.WriteMessage("\nVPFROMREG canceled.");
        return;
      }

      string selectedLayoutName = optionsWindow.SelectedLayoutName;
      ViewportScaleOption selectedScale = ResolveScaleOption(
        optionsWindow.SelectedScaleOption,
        scaleOptions,
        regionWidth,
        regionHeight
      );

      if (selectedScale == null)
      {
        ed.WriteMessage(
          $"\nVPFROMREG: The selected region does not fit within " +
          $"{ViewportAutoFitMaximumWidth:0.##} x {ViewportAutoFitMaximumHeight:0.##} " +
          "paperspace units at any available scale. Select a specific scale instead."
        );
        return;
      }

      double viewportWidth = regionWidth / selectedScale.ModelUnitsPerPaperUnit;
      double viewportHeight = regionHeight / selectedScale.ModelUnitsPerPaperUnit;

      Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(
        "CTAB",
        selectedLayoutName
      );

      PromptPointResult paperSpaceCornerResult = ed.GetPoint(
        new PromptPointOptions("\nSelect the top-right corner for the viewport in paperspace: ")
      );
      if (paperSpaceCornerResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nVPFROMREG canceled.");
        return;
      }

      try
      {
        CreateViewport(
          db,
          selectedLayoutName,
          paperSpaceCornerResult.Value,
          regionExtents,
          regionWidth,
          regionHeight,
          viewportWidth,
          viewportHeight,
          selectedScale
        );

        ed.WriteMessage(
          $"\nVPFROMREG: Created a {selectedScale.DisplayName} viewport on " +
          $"'{selectedLayoutName}' ({viewportWidth:0.##} x {viewportHeight:0.##})."
        );
        ed.Regen();
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nVPFROMREG error: {ex.Message}");
      }
    }

    private static List<string> GetPaperLayoutNames(Database db)
    {
      List<Tuple<int, string>> layouts = new List<Tuple<int, string>>();

      using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
      {
        DBDictionary layoutDictionary = tr.GetObject(
          db.LayoutDictionaryId,
          OpenMode.ForRead
        ) as DBDictionary;

        foreach (DBDictionaryEntry entry in layoutDictionary)
        {
          Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
          if (layout != null && !layout.ModelType)
          {
            layouts.Add(Tuple.Create(layout.TabOrder, layout.LayoutName));
          }
        }

        tr.Commit();
      }

      return layouts
        .OrderBy(layout => layout.Item1)
        .ThenBy(layout => layout.Item2, StringComparer.OrdinalIgnoreCase)
        .Select(layout => layout.Item2)
        .ToList();
    }

    internal static List<ViewportScaleOption> CreateViewportScaleOptions()
    {
      return new List<ViewportScaleOption>
      {
        ViewportScaleOption.AutoFit(),
        ViewportScaleOption.Custom(),
        ViewportScaleOption.Fixed("6\" = 1'-0\"", 2.0),
        ViewportScaleOption.Fixed("3\" = 1'-0\"", 4.0),
        ViewportScaleOption.Fixed("1-1/2\" = 1'-0\"", 8.0),
        ViewportScaleOption.Fixed("1\" = 1'-0\"", 12.0),
        ViewportScaleOption.Fixed("3/4\" = 1'-0\"", 16.0),
        ViewportScaleOption.Fixed("1/2\" = 1'-0\"", 24.0),
        ViewportScaleOption.Fixed("3/8\" = 1'-0\"", 32.0),
        ViewportScaleOption.Fixed("1/4\" = 1'-0\"", 48.0),
        ViewportScaleOption.Fixed("3/16\" = 1'-0\"", 64.0),
        ViewportScaleOption.Fixed("1/8\" = 1'-0\"", 96.0),
        ViewportScaleOption.Fixed("3/32\" = 1'-0\"", 128.0),
        ViewportScaleOption.Fixed("1/16\" = 1'-0\"", 192.0),
        ViewportScaleOption.Fixed("1/32\" = 1'-0\"", 384.0),
        ViewportScaleOption.Fixed("1/64\" = 1'-0\"", 768.0),
        ViewportScaleOption.Fixed("1/128\" = 1'-0\"", 1536.0)
      };
    }

    internal static ViewportScaleOption ResolveScaleOption(
      ViewportScaleOption requestedScale,
      IEnumerable<ViewportScaleOption> scaleOptions,
      double regionWidth,
      double regionHeight
    )
    {
      if (requestedScale == null)
      {
        return null;
      }

      if (!requestedScale.IsAutoFit)
      {
        return requestedScale.IsCustom ? null : requestedScale;
      }

      return scaleOptions
        .Where(scale => !scale.IsAutoFit && !scale.IsCustom)
        .OrderBy(scale => scale.ModelUnitsPerPaperUnit)
        .FirstOrDefault(scale =>
          regionWidth / scale.ModelUnitsPerPaperUnit <= ViewportAutoFitMaximumWidth &&
          regionHeight / scale.ModelUnitsPerPaperUnit <= ViewportAutoFitMaximumHeight
        );
    }

    private static void CreateViewport(
      Database db,
      string layoutName,
      Point3d topRightCorner,
      Extents3d regionExtents,
      double regionWidth,
      double regionHeight,
      double viewportWidth,
      double viewportHeight,
      ViewportScaleOption selectedScale
    )
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        DBDictionary layoutDictionary = tr.GetObject(
          db.LayoutDictionaryId,
          OpenMode.ForRead
        ) as DBDictionary;
        if (!layoutDictionary.Contains(layoutName))
        {
          throw new InvalidOperationException($"Layout '{layoutName}' no longer exists.");
        }

        Layout layout = tr.GetObject(
          layoutDictionary.GetAt(layoutName),
          OpenMode.ForRead
        ) as Layout;
        BlockTableRecord paperSpace = tr.GetObject(
          layout.BlockTableRecordId,
          OpenMode.ForWrite
        ) as BlockTableRecord;

        EnsureDefpointsLayer(db, tr);

        Point2d modelSpaceCenter = new Point2d(
          regionExtents.MinPoint.X + regionWidth / 2.0,
          regionExtents.MinPoint.Y + regionHeight / 2.0
        );

        using (Viewport viewport = new Viewport())
        {
          viewport.CenterPoint = new Point3d(
            topRightCorner.X - viewportWidth / 2.0,
            topRightCorner.Y - viewportHeight / 2.0,
            0.0
          );
          viewport.Width = viewportWidth;
          viewport.Height = viewportHeight;
          viewport.Layer = "DEFPOINTS";
          viewport.ViewTarget = new Point3d(modelSpaceCenter.X, modelSpaceCenter.Y, 0.0);
          viewport.ViewDirection = Vector3d.ZAxis;

          paperSpace.AppendEntity(viewport);
          tr.AddNewlyCreatedDBObject(viewport, true);

          viewport.On = true;
          viewport.CustomScale = selectedScale.CustomScale;
          viewport.Locked = true;
        }

        tr.Commit();
      }
    }

    private static void EnsureDefpointsLayer(Database db, Transaction tr)
    {
      LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
      if (layerTable.Has("DEFPOINTS"))
      {
        return;
      }

      layerTable.UpgradeOpen();
      using (LayerTableRecord layerRecord = new LayerTableRecord
      {
        Name = "DEFPOINTS",
        Color = Color.FromColorIndex(ColorMethod.ByAci, 7)
      })
      {
        layerTable.Add(layerRecord);
        tr.AddNewlyCreatedDBObject(layerRecord, true);
      }
    }

    private static (Point3d Min, Point3d Max) GetCorrectedPoints(Point3d p1, Point3d p2)
    {
      Point3d minPoint = new Point3d(
        Math.Min(p1.X, p2.X),
        Math.Min(p1.Y, p2.Y),
        0.0
      );
      Point3d maxPoint = new Point3d(
        Math.Max(p1.X, p2.X),
        Math.Max(p1.Y, p2.Y),
        0.0
      );

      return (minPoint, maxPoint);
    }
  }

  internal sealed class ViewportScaleOption
  {
    private ViewportScaleOption(
      string displayName,
      double modelUnitsPerPaperUnit,
      bool isAutoFit,
      bool isCustom
    )
    {
      DisplayName = displayName;
      ModelUnitsPerPaperUnit = modelUnitsPerPaperUnit;
      IsAutoFit = isAutoFit;
      IsCustom = isCustom;
    }

    public string DisplayName { get; }
    public double ModelUnitsPerPaperUnit { get; }
    public double CustomScale => IsAutoFit || IsCustom ? 0.0 : 1.0 / ModelUnitsPerPaperUnit;
    public bool IsAutoFit { get; }
    public bool IsCustom { get; }

    public static ViewportScaleOption AutoFit()
    {
      return new ViewportScaleOption("Auto Fit (maximum 30 x 22)", 0.0, true, false);
    }

    public static ViewportScaleOption Custom()
    {
      return new ViewportScaleOption("Custom Scale...", 0.0, false, true);
    }

    public static ViewportScaleOption Fixed(string displayName, double modelUnitsPerPaperUnit)
    {
      if (modelUnitsPerPaperUnit <= 0.0)
      {
        throw new ArgumentOutOfRangeException(nameof(modelUnitsPerPaperUnit));
      }

      return new ViewportScaleOption(displayName, modelUnitsPerPaperUnit, false, false);
    }

    public static bool TryCreateCustom(
      string input,
      out ViewportScaleOption scaleOption,
      out string validationMessage
    )
    {
      scaleOption = null;
      validationMessage = string.Empty;
      string normalized = (input ?? string.Empty)
        .Trim()
        .Replace('\u2033', '"')
        .Replace('\u201D', '"')
        .Replace('\u201C', '"')
        .Replace('\u2032', '\'')
        .Replace('\u2019', '\'');

      if (string.IsNullOrWhiteSpace(normalized))
      {
        validationMessage = "Enter a custom scale.";
        return false;
      }

      double modelUnitsPerPaperUnit;
      if (!TryParseArchitecturalScale(normalized, out modelUnitsPerPaperUnit) &&
          !TryParseRatioScale(normalized, out modelUnitsPerPaperUnit))
      {
        validationMessage =
          "Use a format such as 1\" = 10'-0\", 1:120, or 1/120.";
        return false;
      }

      if (modelUnitsPerPaperUnit <= 0.0 ||
          double.IsNaN(modelUnitsPerPaperUnit) ||
          double.IsInfinity(modelUnitsPerPaperUnit))
      {
        validationMessage = "The custom scale must be greater than zero.";
        return false;
      }

      scaleOption = Fixed($"Custom: {normalized}", modelUnitsPerPaperUnit);
      return true;
    }

    private static bool TryParseArchitecturalScale(
      string input,
      out double modelUnitsPerPaperUnit
    )
    {
      modelUnitsPerPaperUnit = 0.0;
      Match match = Regex.Match(
        input,
        @"^\s*(?<paper>[\d\s./]+?)\s*(?:""|in(?:ch(?:es)?)?)\s*=\s*" +
        @"(?<feet>[\d\s./]+?)\s*(?:'|ft|feet)" +
        @"(?:\s*-\s*(?<inches>[\d\s./]+?)\s*(?:""|in(?:ch(?:es)?)?))?\s*$",
        RegexOptions.IgnoreCase
      );
      if (!match.Success ||
          !TryParseScaleNumber(match.Groups["paper"].Value, out double paperInches) ||
          !TryParseScaleNumber(match.Groups["feet"].Value, out double modelFeet))
      {
        return false;
      }

      double modelInches = 0.0;
      System.Text.RegularExpressions.Group inchesGroup = match.Groups["inches"];
      if (inchesGroup.Success &&
          !TryParseScaleNumber(inchesGroup.Value, out modelInches))
      {
        return false;
      }

      double representedModelInches = modelFeet * 12.0 + modelInches;
      if (paperInches <= 0.0 || representedModelInches <= 0.0)
      {
        return false;
      }

      modelUnitsPerPaperUnit = representedModelInches / paperInches;
      return true;
    }

    private static bool TryParseRatioScale(
      string input,
      out double modelUnitsPerPaperUnit
    )
    {
      modelUnitsPerPaperUnit = 0.0;
      Match match = Regex.Match(
        input,
        @"^\s*(?<paper>\d+(?:\.\d+)?)\s*(?<separator>[:/])\s*" +
        @"(?<model>\d+(?:\.\d+)?)\s*$"
      );
      if (!match.Success ||
          !TryParseScaleNumber(match.Groups["paper"].Value, out double paperUnits) ||
          !TryParseScaleNumber(match.Groups["model"].Value, out double modelUnits) ||
          paperUnits <= 0.0 ||
          modelUnits <= 0.0)
      {
        return false;
      }

      modelUnitsPerPaperUnit = modelUnits / paperUnits;
      return true;
    }

    private static bool TryParseScaleNumber(string input, out double value)
    {
      value = 0.0;
      string normalized = Regex.Replace((input ?? string.Empty).Trim(), @"\s+", " ");
      if (string.IsNullOrEmpty(normalized))
      {
        return false;
      }

      string wholeNumberText = string.Empty;
      string fractionText = normalized;
      int spaceIndex = normalized.IndexOf(' ');
      if (spaceIndex >= 0)
      {
        wholeNumberText = normalized.Substring(0, spaceIndex);
        fractionText = normalized.Substring(spaceIndex + 1);
      }

      double wholeNumber = 0.0;
      if (!string.IsNullOrEmpty(wholeNumberText) &&
          !TryParseDouble(wholeNumberText, out wholeNumber))
      {
        return false;
      }

      if (fractionText.Contains("/"))
      {
        string[] fractionParts = fractionText.Split('/');
        if (fractionParts.Length != 2 ||
            !TryParseDouble(fractionParts[0], out double numerator) ||
            !TryParseDouble(fractionParts[1], out double denominator) ||
            denominator == 0.0)
        {
          return false;
        }

        value = wholeNumber + numerator / denominator;
        return true;
      }

      if (!string.IsNullOrEmpty(wholeNumberText))
      {
        return false;
      }

      return TryParseDouble(fractionText, out value);
    }

    private static bool TryParseDouble(string input, out double value)
    {
      return double.TryParse(
          (input ?? string.Empty).Trim(),
          NumberStyles.Float,
          CultureInfo.InvariantCulture,
          out value
        ) ||
        double.TryParse(
          (input ?? string.Empty).Trim(),
          NumberStyles.Float,
          CultureInfo.CurrentCulture,
          out value
        );
    }

    public override string ToString()
    {
      return DisplayName;
    }
  }
}
