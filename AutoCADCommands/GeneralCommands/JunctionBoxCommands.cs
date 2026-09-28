using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace ElectricalCommands
{
  internal enum JunctionBoxStyle
  {
    Plain,
    Switch,
    WallMount
  }

  public partial class GeneralCommands
  {
    private const string JunctionBoxBlockName = "JBOX";

    // Remembered for the AutoCAD session so repeating JBOX reuses the last choices.
    private static JunctionBoxStyle _lastJunctionBoxStyle = JunctionBoxStyle.Plain;
    private static SymbolSide _lastSwitchSide = SymbolSide.North;
    private static SymbolSide _lastWallMountSide = SymbolSide.East;

    [CommandMethod("JBOX", CommandFlags.Modal)]
    public static void InsertJunctionBox()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage(
          "\nJunction box insertion requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
        return;
      }

      double blockScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
      double rotation = SymbolGeometry.ResolveUcsRotation(ed);

      JunctionBoxStyle style = _lastJunctionBoxStyle;
      Point3d location;
      while (true)
      {
        using (SymbolLocationJig locationJig = new SymbolLocationJig(
          JunctionBoxGeometry.CreateEntities(db, style, GetLastJunctionBoxSide(style)),
          $"\nSpecify location for {DescribeJunctionBoxStyle(style)} " +
            "or [Plain/Switch/Wallmount]: ",
          "Plain Switch Wallmount",
          blockScale,
          rotation))
        {
          PromptResult locationResult = ed.Drag(locationJig);
          if (locationResult.Status == PromptStatus.Keyword)
          {
            if (Enum.TryParse(
              locationResult.StringResult,
              true,
              out JunctionBoxStyle selectedStyle))
            {
              style = selectedStyle;
            }
            continue;
          }
          if (locationResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nJBOX canceled.");
            return;
          }

          location = locationJig.Location;
          break;
        }
      }

      SymbolSide side = GetLastJunctionBoxSide(style);
      if (style != JunctionBoxStyle.Plain)
      {
        string accessoryName = style == JunctionBoxStyle.Switch ? "switch" : "wall-mount";
        if (!TryPromptSymbolSide(
          ed,
          selectedSide => JunctionBoxGeometry.CreateEntities(db, style, selectedSide),
          side,
          $"\nSpecify {accessoryName} side or [North/East/South/West]: ",
          location,
          blockScale,
          rotation,
          JunctionBoxGeometry.Radius * blockScale,
          out side))
        {
          ed.WriteMessage("\nJBOX canceled.");
          return;
        }
      }

      try
      {
        string blockName = ResolveJunctionBoxBlockName(style, side);
        InsertSymbolBlock(
          db,
          blockName,
          () => JunctionBoxGeometry.CreateEntities(db, style, side),
          location,
          blockScale,
          rotation);

        _lastJunctionBoxStyle = style;
        if (style == JunctionBoxStyle.Switch)
        {
          _lastSwitchSide = side;
        }
        else if (style == JunctionBoxStyle.WallMount)
        {
          _lastWallMountSide = side;
        }

        ed.WriteMessage(
          $"\nInserted {blockName} at {scale.DisplayText} " +
          $"(X/Y/Z scale {FormatNumber(blockScale)}).");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to insert the junction box: {ex.Message}");
      }
    }

    // Returns false when the user cancels. Enter accepts the side currently shown.
    private static bool TryPromptSymbolSide(
      Editor editor,
      Func<SymbolSide, List<Entity>> createPreviewEntities,
      SymbolSide initialSide,
      string message,
      Point3d location,
      double blockScale,
      double rotation,
      double minimumPickDistance,
      out SymbolSide side)
    {
      side = initialSide;
      using (SymbolSideJig sideJig = new SymbolSideJig(
        createPreviewEntities,
        initialSide,
        message,
        location,
        blockScale,
        rotation,
        minimumPickDistance))
      {
        PromptResult sideResult = editor.Drag(sideJig);
        if (sideResult.Status == PromptStatus.Keyword)
        {
          if (Enum.TryParse(sideResult.StringResult, true, out SymbolSide selectedSide))
          {
            side = selectedSide;
          }
          return true;
        }
        if (sideResult.Status == PromptStatus.OK ||
            sideResult.Status == PromptStatus.None)
        {
          side = sideJig.Side;
          return true;
        }
        return false;
      }
    }

    private static void InsertSymbolBlock(
      Database database,
      string blockName,
      Func<List<Entity>> createEntities,
      Point3d location,
      double blockScale,
      double rotation)
    {
      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        ObjectId blockDefinitionId = SymbolGeometry.EnsureBlock(
          transaction,
          database,
          blockName,
          createEntities);

        BlockTableRecord currentSpace = (BlockTableRecord)transaction.GetObject(
          database.CurrentSpaceId,
          OpenMode.ForWrite);

        BlockReference blockReference = new BlockReference(
          location,
          blockDefinitionId);
        blockReference.SetDatabaseDefaults(database);
        blockReference.ScaleFactors = new Scale3d(blockScale);
        blockReference.Rotation = rotation;

        currentSpace.AppendEntity(blockReference);
        transaction.AddNewlyCreatedDBObject(blockReference, true);

        transaction.Commit();
      }
    }

    private static SymbolSide GetLastJunctionBoxSide(JunctionBoxStyle style)
    {
      switch (style)
      {
        case JunctionBoxStyle.Switch:
          return _lastSwitchSide;
        case JunctionBoxStyle.WallMount:
          return _lastWallMountSide;
        default:
          return SymbolSide.North;
      }
    }

    private static string DescribeJunctionBoxStyle(JunctionBoxStyle style)
    {
      switch (style)
      {
        case JunctionBoxStyle.Switch:
          return "junction box with switch";
        case JunctionBoxStyle.WallMount:
          return "wall-mount junction box";
        default:
          return "junction box";
      }
    }

    private static string ResolveJunctionBoxBlockName(
      JunctionBoxStyle style,
      SymbolSide side)
    {
      char sideLetter = side.ToString()[0];
      switch (style)
      {
        case JunctionBoxStyle.Switch:
          return $"{JunctionBoxBlockName}-SWITCH-{sideLetter}";
        case JunctionBoxStyle.WallMount:
          return $"{JunctionBoxBlockName}-WALLMOUNT-{sideLetter}";
        default:
          return JunctionBoxBlockName;
      }
    }
  }

  // Builds the junction box symbol around the origin in plotted inches.
  internal static class JunctionBoxGeometry
  {
    // 5.3865" model diameter at 1/4" = 1'-0".
    internal const double Radius =
      5.3865 / SymbolGeometry.QuarterScaleModelInchesPerPlottedInch / 2.0;

    internal static List<Entity> CreateEntities(
      Database database,
      JunctionBoxStyle style,
      SymbolSide side
    )
    {
      List<Entity> entities = new List<Entity>
      {
        SymbolGeometry.Prepare(
          database,
          new Circle(Point3d.Origin, Vector3d.ZAxis, Radius),
          SymbolGeometry.YellowColorIndex
        ),
        SymbolGeometry.Prepare(database, CreateLetterJ(), SymbolGeometry.WhiteColorIndex),
      };

      if (style == JunctionBoxStyle.Plain)
      {
        return entities;
      }

      // Accessories are drawn on the north side and rotated to the requested side;
      // the J itself always stays upright.
      Matrix3d toSide = Matrix3d.Rotation(
        SymbolGeometry.ResolveSideRotation(side),
        Vector3d.ZAxis,
        Point3d.Origin
      );
      bool isSwitch = style == JunctionBoxStyle.Switch;
      IEnumerable<Entity> accessory = isSwitch ? CreateSwitchLeg() : CreateWallMount();
      foreach (Entity entity in accessory)
      {
        entity.TransformBy(toSide);
        entities.Add(SymbolGeometry.Prepare(
          database,
          entity,
          isSwitch ? SymbolGeometry.RedColorIndex : SymbolGeometry.YellowColorIndex
        ));
      }
      return entities;
    }

    private static Polyline CreateLetterJ()
    {
      Polyline letter = new Polyline();
      letter.AddVertexAt(0, new Point2d(0.25 * Radius, 0.55 * Radius), 0.0, 0.0, 0.0);
      // Bulge -1 is a clockwise half circle, forming the hook through the bottom.
      letter.AddVertexAt(1, new Point2d(0.25 * Radius, -0.30 * Radius), -1.0, 0.0, 0.0);
      letter.AddVertexAt(2, new Point2d(-0.25 * Radius, -0.30 * Radius), 0.0, 0.0, 0.0);
      letter.AddVertexAt(3, new Point2d(-0.25 * Radius, -0.18 * Radius), 0.0, 0.0, 0.0);
      return letter;
    }

    private static IEnumerable<Entity> CreateSwitchLeg()
    {
      double bowlRadius = 0.4 * Radius;
      yield return new Line(
        new Point3d(0.0, Radius, 0.0),
        new Point3d(0.0, 3.4 * Radius, 0.0)
      );
      // Upper bowl of the S: from the right side, over the top, to the middle.
      yield return new Arc(
        new Point3d(0.0, 2.6 * Radius, 0.0),
        bowlRadius,
        0.0,
        1.5 * Math.PI
      );
      // Lower bowl: from the lower left, under the bottom and up to the middle.
      yield return new Arc(
        new Point3d(0.0, 1.8 * Radius, 0.0),
        bowlRadius,
        Math.PI,
        0.5 * Math.PI
      );
    }

    private static IEnumerable<Entity> CreateWallMount()
    {
      double wallOffset = 1.7 * Radius;
      yield return new Line(
        new Point3d(0.0, Radius, 0.0),
        new Point3d(0.0, wallOffset, 0.0)
      );
      yield return new Line(
        new Point3d(-0.55 * Radius, wallOffset, 0.0),
        new Point3d(0.55 * Radius, wallOffset, 0.0)
      );
    }
  }
}
