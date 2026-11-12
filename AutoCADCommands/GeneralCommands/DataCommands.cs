using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    // Avoid reusing imported dynamic DATA blocks, whose default angle/visibility
    // can make the inserted symbol disagree with our static preview.
    private const string DataBlockName = "ACIES_DATA";

    // Remembered for the AutoCAD session so repeating DATA reuses the last direction.
    private static SymbolSide _lastDataSide = SymbolSide.South;

    [CommandMethod("DATA", CommandFlags.Modal)]
    public static void InsertData()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage(
          "\nData symbol insertion requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
        return;
      }

      double blockScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
      double ucsRotation = SymbolGeometry.ResolveUcsRotation(ed);
      SymbolSide side = _lastDataSide;

      Point3d location;
      using (SymbolLocationJig locationJig = new SymbolLocationJig(
        DataGeometry.CreateEntities(db, side),
        "\nSpecify location for data symbol: ",
        null,
        blockScale,
        ucsRotation))
      {
        if (ed.Drag(locationJig).Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nDATA canceled.");
          return;
        }
        location = locationJig.Location;
      }

      if (!TryPromptSymbolSide(
        ed,
        selectedSide => DataGeometry.CreateEntities(db, selectedSide),
        side,
        "\nSpecify data symbol point direction or [North/East/South/West]: ",
        location,
        blockScale,
        ucsRotation,
        DataGeometry.HalfHeight * blockScale,
        out side))
      {
        ed.WriteMessage("\nDATA canceled.");
        return;
      }

      try
      {
        // Zero rotation points north; a 180-degree rotation points south.
        InsertSymbolBlock(
          db,
          DataBlockName,
          () => DataGeometry.CreateEntities(db, SymbolSide.North),
          location,
          blockScale,
          ucsRotation + DataGeometry.ResolveRotation(side));

        _lastDataSide = side;
        ed.WriteMessage(
          $"\nInserted DATA pointing {side.ToString().ToLowerInvariant()} " +
          $"at {scale.DisplayText} (X/Y/Z scale {FormatNumber(blockScale)}).");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to insert the data symbol: {ex.Message}");
      }
    }
  }

  // Builds the data symbol centered on the origin in plotted inches:
  // a cyan triangle whose apex points toward the requested side. The block
  // definition is drawn pointing north, so 0/90/180/270 face N/W/S/E.
  internal static class DataGeometry
  {
    private const double ModelToPlotted = SymbolGeometry.QuarterScaleModelInchesPerPlottedInch;

    // Model inches at 1/4" = 1'-0".
    internal const double HalfWidth = 5.3865 / ModelToPlotted / 2.0;
    internal const double HalfHeight = 5.3865 / ModelToPlotted / 2.0;

    // Rotation that turns the north-pointing definition toward the given side.
    internal static double ResolveRotation(SymbolSide side)
    {
      switch (side)
      {
        case SymbolSide.North:
          return 0.0;
        case SymbolSide.East:
          return 3.0 * Math.PI / 2.0;
        case SymbolSide.West:
          return Math.PI / 2.0;
        default:
          return Math.PI;
      }
    }

    internal static List<Entity> CreateEntities(Database database, SymbolSide side)
    {
      Polyline triangle = new Polyline();
      triangle.AddVertexAt(0, new Point2d(-HalfWidth, -HalfHeight), 0.0, 0.0, 0.0);
      triangle.AddVertexAt(1, new Point2d(HalfWidth, -HalfHeight), 0.0, 0.0, 0.0);
      triangle.AddVertexAt(2, new Point2d(0.0, HalfHeight), 0.0, 0.0, 0.0);
      triangle.Closed = true;
      triangle.TransformBy(Matrix3d.Rotation(
        ResolveRotation(side),
        Vector3d.ZAxis,
        Point3d.Origin));

      return new List<Entity>
      {
        SymbolGeometry.Prepare(database, triangle, SymbolGeometry.CyanColorIndex),
      };
    }
  }
}
