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
    private const string DisconnectBlockName = "DISCONNECT";

    // Remembered for the AutoCAD session so repeating DISCONNECT reuses the last side.
    private static SymbolSide _lastDisconnectSide = SymbolSide.North;

    [CommandMethod("DISCONNECT", CommandFlags.Modal)]
    public static void InsertDisconnect()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage(
          "\nDisconnect insertion requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
        return;
      }

      double blockScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
      double ucsRotation = SymbolGeometry.ResolveUcsRotation(ed);
      SymbolSide side = _lastDisconnectSide;

      Point3d location;
      using (SymbolLocationJig locationJig = new SymbolLocationJig(
        DisconnectGeometry.CreateEntities(db, side),
        "\nSpecify location for disconnect: ",
        null,
        blockScale,
        ucsRotation))
      {
        if (ed.Drag(locationJig).Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nDISCONNECT canceled.");
          return;
        }
        location = locationJig.Location;
      }

      if (!TryPromptSymbolSide(
        ed,
        selectedSide => DisconnectGeometry.CreateEntities(db, selectedSide),
        side,
        "\nSpecify disconnect handle side or [North/East/South/West]: ",
        location,
        blockScale,
        ucsRotation,
        DisconnectGeometry.HalfHeight * blockScale,
        out side))
      {
        ed.WriteMessage("\nDISCONNECT canceled.");
        return;
      }

      try
      {
        // One definition drawn with the handle north; the insert is rotated to the side.
        InsertSymbolBlock(
          db,
          DisconnectBlockName,
          () => DisconnectGeometry.CreateEntities(db, SymbolSide.North),
          location,
          blockScale,
          ucsRotation + SymbolGeometry.ResolveSideRotation(side));

        _lastDisconnectSide = side;
        ed.WriteMessage(
          $"\nInserted {DisconnectBlockName} with the handle {side.ToString().ToLowerInvariant()} " +
          $"at {scale.DisplayText} (X/Y/Z scale {FormatNumber(blockScale)}).");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to insert the disconnect: {ex.Message}");
      }
    }
  }

  // Builds the disconnect symbol centered on the origin in plotted inches:
  // a rectangle with the handle rising from the middle of the north edge
  // and turning east.
  internal static class DisconnectGeometry
  {
    private const double ModelToPlotted = SymbolGeometry.QuarterScaleModelInchesPerPlottedInch;

    // Model inches at 1/4" = 1'-0"; the width matches the JBOX diameter.
    internal const double HalfWidth = 5.3865 / ModelToPlotted / 2.0;
    internal const double HalfHeight = 7.4394 / ModelToPlotted / 2.0;
    private const double StemLength = 2.6018 / ModelToPlotted;
    private const double HandleLength = 2.2969 / ModelToPlotted;

    internal static List<Entity> CreateEntities(Database database, SymbolSide side)
    {
      Polyline body = new Polyline();
      body.AddVertexAt(0, new Point2d(-HalfWidth, -HalfHeight), 0.0, 0.0, 0.0);
      body.AddVertexAt(1, new Point2d(HalfWidth, -HalfHeight), 0.0, 0.0, 0.0);
      body.AddVertexAt(2, new Point2d(HalfWidth, HalfHeight), 0.0, 0.0, 0.0);
      body.AddVertexAt(3, new Point2d(-HalfWidth, HalfHeight), 0.0, 0.0, 0.0);
      body.Closed = true;

      Polyline handle = new Polyline();
      handle.AddVertexAt(0, new Point2d(0.0, HalfHeight), 0.0, 0.0, 0.0);
      handle.AddVertexAt(1, new Point2d(0.0, HalfHeight + StemLength), 0.0, 0.0, 0.0);
      handle.AddVertexAt(2, new Point2d(HandleLength, HalfHeight + StemLength), 0.0, 0.0, 0.0);

      Matrix3d toSide = Matrix3d.Rotation(
        SymbolGeometry.ResolveSideRotation(side),
        Vector3d.ZAxis,
        Point3d.Origin);
      List<Entity> entities = new List<Entity>();
      foreach (Entity entity in new Entity[] { body, handle })
      {
        entity.TransformBy(toSide);
        entities.Add(SymbolGeometry.Prepare(database, entity, SymbolGeometry.RedColorIndex));
      }
      return entities;
    }
  }
}
