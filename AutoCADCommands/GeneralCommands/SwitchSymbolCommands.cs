using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace ElectricalCommands
{
  internal enum SwitchSymbolStyle
  {
    Standard,
    Dimmer,
    Occupancy
  }

  public partial class GeneralCommands
  {
    // Remembered for the AutoCAD session so repeating SWITCH reuses the last style.
    private static SwitchSymbolStyle _lastSwitchSymbolStyle = SwitchSymbolStyle.Standard;

    [CommandMethod("SWITCH", CommandFlags.Modal)]
    public static void InsertSwitchSymbol()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage(
          "\nSwitch insertion requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
        return;
      }

      double blockScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
      double rotation = SymbolGeometry.ResolveUcsRotation(ed);

      SwitchSymbolStyle style = _lastSwitchSymbolStyle;
      Point3d location;
      while (true)
      {
        using (SymbolLocationJig locationJig = new SymbolLocationJig(
          SwitchSymbolGeometry.CreateEntities(db, style),
          $"\nSpecify location for {style.ToString().ToLowerInvariant()} switch " +
            "or [Standard/Dimmer/Occupancy]: ",
          "Standard Dimmer Occupancy",
          blockScale,
          rotation))
        {
          PromptResult locationResult = ed.Drag(locationJig);
          if (locationResult.Status == PromptStatus.Keyword)
          {
            if (Enum.TryParse(
              locationResult.StringResult,
              true,
              out SwitchSymbolStyle selectedStyle))
            {
              style = selectedStyle;
            }
            continue;
          }
          if (locationResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nSWITCH canceled.");
            return;
          }

          location = locationJig.Location;
          break;
        }
      }

      try
      {
        string blockName = ResolveSwitchSymbolBlockName(style);
        InsertSymbolBlock(
          db,
          blockName,
          () => SwitchSymbolGeometry.CreateEntities(db, style),
          location,
          blockScale,
          rotation);

        _lastSwitchSymbolStyle = style;
        ed.WriteMessage(
          $"\nInserted {blockName} at {scale.DisplayText} " +
          $"(X/Y/Z scale {FormatNumber(blockScale)}).");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to insert the switch: {ex.Message}");
      }
    }

    private static string ResolveSwitchSymbolBlockName(SwitchSymbolStyle style)
    {
      switch (style)
      {
        case SwitchSymbolStyle.Dimmer:
          return "SWITCH-DIMMER";
        case SwitchSymbolStyle.Occupancy:
          return "SWITCH-OCCUPANCY";
        default:
          return "SWITCH";
      }
    }
  }

  // Builds the switch symbol in plotted inches: a "$" centered on the origin,
  // followed by "D" (dimmer) or "OS" (occupancy sensor). Glyphs are drawn as
  // polylines in units of the letter height.
  internal static class SwitchSymbolGeometry
  {
    // Letter height is 4.5" model at 1/4" = 1'-0", matching SCALEDTEXT.
    private const double LetterHeight =
      4.5 / SymbolGeometry.QuarterScaleModelInchesPerPlottedInch;

    // How far the "$" stroke runs past the top and bottom of the S.
    private const double StrokeHalfLength = 0.92;
    private const double LetterGap = 0.15;

    private static readonly double[] SGlyph =
    {
      0.35, 0.36, 0.15, 0.50, -0.15, 0.50, -0.35, 0.28, -0.35, 0.20, -0.20, 0.03,
      0.20, -0.03, 0.35, -0.20, 0.35, -0.28, 0.15, -0.50, -0.15, -0.50, -0.35, -0.36,
    };

    private static readonly double[] DGlyph =
    {
      0.0, 0.5, 0.20, 0.5, 0.38, 0.42, 0.45, 0.25, 0.45, -0.25, 0.38, -0.42,
      0.20, -0.5, 0.0, -0.5,
    };

    private static readonly double[] OGlyph =
    {
      -0.18, 0.5, 0.18, 0.5, 0.33, 0.36, 0.39, 0.12, 0.39, -0.12, 0.33, -0.36,
      0.18, -0.5, -0.18, -0.5, -0.33, -0.36, -0.39, -0.12, -0.39, 0.12, -0.33, 0.36,
    };

    internal static List<Entity> CreateEntities(Database database, SwitchSymbolStyle style)
    {
      List<Entity> entities = new List<Entity>
      {
        Prepare(database, new Line(
          new Point3d(0.0, -StrokeHalfLength * LetterHeight, 0.0),
          new Point3d(0.0, StrokeHalfLength * LetterHeight, 0.0))),
        Prepare(database, CreateGlyph(SGlyph, 0.0, false)),
      };

      // Glyph extents to the right of the "$": S is +/-0.35, D spans 0..0.45, O is +/-0.39.
      double cursor = 0.35 + LetterGap;
      if (style == SwitchSymbolStyle.Dimmer)
      {
        entities.Add(Prepare(database, CreateGlyph(DGlyph, cursor, true)));
      }
      else if (style == SwitchSymbolStyle.Occupancy)
      {
        entities.Add(Prepare(database, CreateGlyph(OGlyph, cursor + 0.39, true)));
        cursor += 0.78 + LetterGap;
        entities.Add(Prepare(database, CreateGlyph(SGlyph, cursor + 0.35, false)));
      }
      return entities;
    }

    private static Entity Prepare(Database database, Entity entity)
    {
      return SymbolGeometry.Prepare(database, entity, SymbolGeometry.CyanColorIndex);
    }

    private static Polyline CreateGlyph(double[] points, double offsetX, bool closed)
    {
      Polyline glyph = new Polyline();
      for (int i = 0; i < points.Length / 2; i++)
      {
        glyph.AddVertexAt(
          i,
          new Point2d(
            (points[2 * i] + offsetX) * LetterHeight,
            points[2 * i + 1] * LetterHeight),
          0.0, 0.0, 0.0);
      }
      glyph.Closed = closed;
      return glyph;
    }
  }
}
