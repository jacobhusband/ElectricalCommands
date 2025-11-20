using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("CREATEVIEWPORTFROMREGION")]
    public void CREATEVIEWPORTFROMREGION()
    {
      Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      // --- 1. Select the region in Modelspace FIRST ---
      PromptPointOptions pointOpts1 = new PromptPointOptions(
          "\nPlease select the first corner of the region in modelspace:"
      );
      PromptPointResult pointResult1 = ed.GetPoint(pointOpts1);
      if (pointResult1.Status != PromptStatus.OK) return;

      PromptPointOptions pointOpts2 = new PromptPointOptions(
          "\nPlease select the opposite corner of the region in modelspace:"
      );
      pointOpts2.BasePoint = pointResult1.Value;
      pointOpts2.UseBasePoint = true;
      PromptPointResult pointResult2 = ed.GetPoint(pointOpts2);
      if (pointResult2.Status != PromptStatus.OK) return;

      var correctedPoints = GetCorrectedPoints(pointResult1.Value, pointResult2.Value);
      Extents3d rectExtents = new Extents3d(correctedPoints.Min, correctedPoints.Max);
      double rectWidth = rectExtents.MaxPoint.X - rectExtents.MinPoint.X;
      double rectHeight = rectExtents.MaxPoint.Y - rectExtents.MinPoint.Y;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        // --- 2. Determine the target Paperspace Layout ---
        string matchedLayoutName = null;
        string userInputForError = "";

        DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
        List<string> paperLayouts = new List<string>();

        foreach (DBDictionaryEntry entry in layoutDict)
        {
          if (!entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
          {
            paperLayouts.Add(entry.Key);
          }
        }

        if (paperLayouts.Count == 1)
        {
          matchedLayoutName = paperLayouts[0];
          ed.WriteMessage($"\nAutomatically selecting the only available paperspace layout: {matchedLayoutName}");
        }
        else if (paperLayouts.Count > 1)
        {
          // CHANGED: Use PromptStringOptions to allow spaces in layout names (e.g. "Layout 1")
          PromptStringOptions pso = new PromptStringOptions("\nPlease enter the Layout name: ");
          pso.AllowSpaces = true;

          PromptResult sheetNameResult = ed.GetString(pso);
          if (sheetNameResult.Status != PromptStatus.OK)
          {
            tr.Abort();
            return;
          }

          userInputForError = sheetNameResult.StringResult.Trim(); // Trim accidental leading/trailing spaces

          // CHANGED: Exact match only, case insensitive. No "E" prefix prediction.
          matchedLayoutName = paperLayouts.FirstOrDefault(name =>
              name.Equals(userInputForError, StringComparison.OrdinalIgnoreCase)
          );
        }
        else
        {
          ed.WriteMessage("\nNo paperspace layouts found in the drawing.");
          tr.Abort();
          return;
        }

        if (string.IsNullOrEmpty(matchedLayoutName))
        {
          ed.WriteMessage($"\nNo matching layout found for '{userInputForError}'.");
          tr.Abort();
          return;
        }

        // --- 3. Switch to the selected Paperspace ---
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CTAB", matchedLayoutName);

        // --- 4. Prompt user to select the placement point in Paperspace ---
        PromptPointOptions paperSpacePointOpts = new PromptPointOptions(
            "\nPlease select the top-right corner for the viewport in paperspace:"
        );
        PromptPointResult paperSpaceCornerResult = ed.GetPoint(paperSpacePointOpts);
        if (paperSpaceCornerResult.Status != PromptStatus.OK)
        {
          tr.Abort();
          return;
        }
        Point3d topRightCorner = paperSpaceCornerResult.Value;

        // --- START: Viewport Creation Logic ---
        Dictionary<double, double> scales = new Dictionary<double, double>
                {
                    { 0.25, 48.0 },
                    { 3.0 / 16.0, 64.0 },
                    { 1.0 / 8.0, 96.0 },
                    { 3.0 / 32.0, 128.0 },
                    { 1.0 / 16.0, 192.0 }
                };

        double scaleToFit = 0.0;
        double viewportWidth = 0.0;
        double viewportHeight = 0.0;

        PromptResult result = ed.GetString("\nEnter scale (e.g., 1/4, 3/16) or press Enter to autoscale: ");
        if (result.Status != PromptStatus.OK && result.Status != PromptStatus.None)
        {
          tr.Abort();
          return;
        }

        string input = result.StringResult.Trim();

        if (string.IsNullOrEmpty(input))
        {
          // Autoscaling logic
          foreach (var scaleEntry in scales.OrderByDescending(e => e.Key))
          {
            viewportWidth = rectWidth / scaleEntry.Value;
            viewportHeight = rectHeight / scaleEntry.Value;

            if (viewportWidth <= 30 && viewportHeight <= 22)
            {
              scaleToFit = scaleEntry.Key;
              break;
            }
          }
          if (scaleToFit == 0.0)
          {
            ed.WriteMessage("Couldn't fit the rectangle in the specified scales");
            tr.Abort();
            return;
          }
        }
        else
        {
          // Manual scale logic
          string[] fraction = input.Split('/');
          if (fraction.Length == 2 && double.TryParse(fraction[0], out double numerator) && double.TryParse(fraction[1], out double denominator) && denominator != 0)
          {
            double inputScale = numerator / denominator;
            scaleToFit = scales.Keys.OrderBy(s => Math.Abs(s - inputScale)).First();
            ed.WriteMessage($"\nUsing closest available scale to input: {ScaleToFraction(scaleToFit)}\" = 1'-0\"");

            double scaleFactor = scales[scaleToFit];
            viewportWidth = rectWidth / scaleFactor;
            viewportHeight = rectHeight / scaleFactor;
          }
          else
          {
            ed.WriteMessage("\nInvalid scale format. Please use format like '1/4'.");
            tr.Abort();
            return;
          }
        }

        ObjectId layoutId = layoutDict.GetAt(matchedLayoutName);
        Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

        BlockTableRecord paperSpace = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

        LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

        if (!layerTable.Has("DEFPOINTS"))
        {
          LayerTableRecord layerRecord = new LayerTableRecord
          {
            Name = "DEFPOINTS",
            Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7)
          };
          layerTable.UpgradeOpen();
          layerTable.Add(layerRecord);
          tr.AddNewlyCreatedDBObject(layerRecord, true);
        }

        Point2d modelSpaceCenter = new Point2d(
            (rectExtents.MinPoint.X + rectWidth / 2),
            (rectExtents.MinPoint.Y + rectHeight / 2)
        );

        Viewport viewport = new Viewport();

        // Calculate CenterPoint from user's top-right corner selection
        viewport.CenterPoint = new Point3d(
            topRightCorner.X - (viewportWidth / 2.0),
            topRightCorner.Y - (viewportHeight / 2.0),
            0.0
        );

        viewport.Width = viewportWidth;
        viewport.Height = viewportHeight;
        viewport.CustomScale = scales.ContainsKey(scaleToFit) ? 1.0 / scales[scaleToFit] : 1.0;
        viewport.Layer = "DEFPOINTS";
        viewport.ViewTarget = new Point3d(modelSpaceCenter.X, modelSpaceCenter.Y, 0.0);
        viewport.ViewDirection = new Vector3d(0, 0, 1);

        ed.WriteMessage($"\nSet viewport scale to {ScaleToFraction(12 * viewport.CustomScale)}\" = 1'-0\"");

        paperSpace.AppendEntity(viewport);
        tr.AddNewlyCreatedDBObject(viewport, true);

        viewport.On = true;
        viewport.Locked = true;

        tr.Commit();
      }
      ed.Regen();
    }

    private static string ScaleToFraction(double scale)
    {
      var knownScales = new Dictionary<double, string>
            {
                { 0.25, "1/4" },
                { 3.0 / 16.0, "3/16" },
                { 1.0 / 8.0, "1/8" },
                { 3.0 / 32.0, "3/32" },
                { 0.0625, "1/16" }
            };

      return knownScales.ContainsKey(scale) ? knownScales[scale] : scale.ToString();
    }

    private (Point3d Min, Point3d Max) GetCorrectedPoints(Point3d p1, Point3d p2)
    {
      Point3d minPoint = new Point3d(Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y), 0);

      Point3d maxPoint = new Point3d(Math.Max(p1.X, p2.X), Math.Max(p1.Y, p2.Y), 0);

      return (minPoint, maxPoint);
    }
  }
}