using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
  internal enum KeyNoteLeaderEnd
  {
    Arrow,
    Circle
  }

  public partial class GeneralCommands
  {
    private const string KnLeaderCircleBlockName = "KN_LEADER_CIRCLE";
    private static string _lastKnlValue = "1";
    private static KeyNoteLeaderEnd _lastKnlEnd = KeyNoteLeaderEnd.Arrow;

    [CommandMethod("KNL", CommandFlags.Modal)]
    public void KeyedNoteLeaderCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!TryResolveKnPlacementScale(ed, db, "KNL", out double scaleDenom)) return;

      string keyValue = PromptKnKeyValue(ed, _lastKnlValue);
      if (string.IsNullOrEmpty(keyValue))
      {
        ed.WriteMessage("\nKNL canceled.");
        return;
      }

      ObjectId blockDefId;
      try
      {
        EnsureKnLayer(db);
        ObjectId textStyleId = EnsureKnTextStyle(db);
        if (!EnsureKnBlockDefinition(db, textStyleId))
        {
          ed.WriteMessage($"\nFailed to prepare canonical {KnBlockName} block definition.");
          return;
        }

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
          blockDefId = bt[KnBlockName];
          tr.Commit();
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nKNL setup error: {ex.Message}");
        return;
      }

      Point3d tip;
      while (true)
      {
        PromptPointOptions tipOptions = new PromptPointOptions(string.Empty)
        {
          AllowNone = false
        };
        tipOptions.SetMessageAndKeywords(
          $"\nSpecify leader tip or [Arrow/Circle] (current: {_lastKnlEnd}): ",
          "Arrow Circle");

        PromptPointResult tipResult = ed.GetPoint(tipOptions);
        if (tipResult.Status == PromptStatus.Keyword)
        {
          _lastKnlEnd = KeyNoteLeaderJig.ParseEnd(tipResult.StringResult, _lastKnlEnd);
          continue;
        }
        if (tipResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nKNL canceled.");
          return;
        }
        tip = tipResult.Value;
        break;
      }

      KeyNoteLeaderLayout layout;
      KeyNoteLeaderEnd leaderEnd;
      try
      {
        BlockReference previewNote = new BlockReference(tip, blockDefId)
        {
          Layer = KnLayerName,
          ScaleFactors = new Scale3d(scaleDenom),
          Rotation = 0.0
        };
        using (KeyNoteLeaderJig jig = new KeyNoteLeaderJig(db, previewNote, tip, scaleDenom, _lastKnlEnd))
        {
          PromptResult jigResult;
          while (true)
          {
            jigResult = ed.Drag(jig);
            if (jigResult.Status == PromptStatus.Keyword)
            {
              jig.ApplyKeyword(jigResult.StringResult);
              continue;
            }
            break;
          }

          if (jigResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nKNL canceled.");
            return;
          }
          layout = jig.Layout;
          leaderEnd = jig.End;
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nKNL jig error: {ex.Message}");
        return;
      }

      try
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

          BlockReference br = new BlockReference(layout.NoteCenter, blockDefId)
          {
            Layer = KnLayerName,
            ScaleFactors = new Scale3d(scaleDenom),
            Rotation = 0.0
          };
          space.AppendEntity(br);
          tr.AddNewlyCreatedDBObject(br, true);
          AddKnAttributes(tr, br, blockDefId, keyValue);

          if (layout.HasLeader)
          {
            Leader leader = CreateKnLeader(db, layout, leaderEnd, scaleDenom);
            space.AppendEntity(leader);
            tr.AddNewlyCreatedDBObject(leader, true);
          }

          tr.Commit();
        }

        _lastKnlValue = keyValue;
        _lastKnlEnd = leaderEnd;
        ed.WriteMessage(
          layout.HasLeader
            ? $"\nInserted {KnBlockName} with value {keyValue} and a leader ending in {(leaderEnd == KeyNoteLeaderEnd.Circle ? "a circle" : "an arrow")} on layer {KnLayerName}."
            : $"\nInserted {KnBlockName} with value {keyValue} on layer {KnLayerName}; the note covers the leader tip, so no leader was drawn.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nKNL placement error: {ex.Message}");
      }
    }

    private static bool TryResolveKnPlacementScale(
      Editor ed,
      Database db,
      string commandName,
      out double scaleDenom)
    {
      scaleDenom = 1.0;
      if (db.TileMode)
      {
        ed.WriteMessage($"\n{commandName} requires a paperspace layout. Switch to a layout tab and run again.");
        return false;
      }

      bool inViewportEditing =
        System.Convert.ToInt16(Application.GetSystemVariable("CVPORT")) > 1;
      if (!inViewportEditing)
      {
        ed.WriteMessage("\nPlacing in paperspace at 1:1 (BR scale factor: 1).");
        return true;
      }

      scaleDenom = ResolveViewportScaleDenominator(ed, db);
      if (scaleDenom <= 0.0)
      {
        ed.WriteMessage($"\n{commandName} canceled: Could not resolve active viewport scale.");
        return false;
      }
      string ratioScale = FormatRatio(scaleDenom);
      TrySetCannoscale(ratioScale);
      ed.WriteMessage($"\nUsing viewport scale: {ratioScale} (BR scale factor: {scaleDenom})");
      return true;
    }

    // Fills the block reference's attributes from the definition, as KN does.
    private static void AddKnAttributes(
      Transaction tr,
      BlockReference br,
      ObjectId blockDefId,
      string keyValue)
    {
      BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockDefId, OpenMode.ForRead);
      foreach (ObjectId id in btr)
      {
        AttributeDefinition ad = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
        if (ad == null || ad.Constant) continue;

        AttributeReference ar = new AttributeReference();
        ar.SetAttributeFromBlock(ad, br.BlockTransform);
        ar.TextString = keyValue;
        br.AttributeCollection.AppendAttribute(ar);
        tr.AddNewlyCreatedDBObject(ar, true);
      }
    }

    private static Leader CreateKnLeader(
      Database db,
      KeyNoteLeaderLayout layout,
      KeyNoteLeaderEnd end,
      double scale)
    {
      Leader leader = new Leader();
      leader.SetDatabaseDefaults(db);
      leader.Layer = KnLayerName;
      leader.ColorIndex = 256;
      leader.IsSplined = false;
      leader.HasArrowHead = true;
      // Absolute size so the drawing's DIMSCALE does not change the arrowhead.
      leader.Dimscale = 1.0;
      leader.Dimasz = KeyNoteLeaderLayout.ArrowSize * scale;
      leader.Dimldrblk = end == KeyNoteLeaderEnd.Circle
        ? EnsureKnLeaderCircleBlock(db)
        : ObjectId.Null;
      foreach (Point3d vertex in layout.Path)
      {
        leader.AppendVertex(vertex);
      }
      return leader;
    }

    // Leader arrowhead blocks are drawn at unit size with the leader coming in
    // along -X. AutoCAD's own _DotBlank stops the line at the circle's edge; this
    // one carries the line on to the circle's center.
    private static ObjectId EnsureKnLeaderCircleBlock(Database db)
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        if (bt.Has(KnLeaderCircleBlockName))
        {
          ObjectId existingId = bt[KnLeaderCircleBlockName];
          tr.Commit();
          return existingId;
        }

        bt.UpgradeOpen();
        BlockTableRecord btr = new BlockTableRecord
        {
          Name = KnLeaderCircleBlockName,
          Origin = Point3d.Origin
        };
        ObjectId id = bt.Add(btr);
        tr.AddNewlyCreatedDBObject(btr, true);

        Entity[] parts =
        {
          new Circle(Point3d.Origin, Vector3d.ZAxis, 0.5),
          new Line(Point3d.Origin, new Point3d(-1.0, 0.0, 0.0))
        };
        foreach (Entity part in parts)
        {
          btr.AppendEntity(part);
          tr.AddNewlyCreatedDBObject(part, true);
          part.LayerId = db.LayerZero;
          // ByBlock, so the arrowhead takes the leader's own color.
          part.ColorIndex = 0;
        }

        tr.Commit();
        return id;
      }
    }
  }

  // Where the keyed note sits and how its leader runs, in drawing units.
  // The leader leaves the hexagon's left/right vertex, runs level to the tip's
  // X, then turns toward the tip; when the tip is directly above or below the
  // note it runs straight out of the top or bottom edge instead.
  internal readonly struct KeyNoteLeaderLayout
  {
    // Plotted inches, the same units as the KEYED_NOTE block definition.
    internal const double HexHalfWidth = 0.1252;
    internal const double HexHalfHeight = 0.108426;
    internal const double ArrowSize = 0.09375;

    internal KeyNoteLeaderLayout(Point3d noteCenter, Point3d[] path)
    {
      NoteCenter = noteCenter;
      Path = path;
    }

    internal Point3d NoteCenter { get; }

    // Tip first, ending on the note's edge; empty when the note covers the tip.
    internal Point3d[] Path { get; }

    internal bool HasLeader => Path != null && Path.Length >= 2;

    internal static KeyNoteLeaderLayout Compute(Point3d tip, Point3d requestedCenter, double scale)
    {
      double halfWidth = HexHalfWidth * scale;
      double halfHeight = HexHalfHeight * scale;
      // A vertical leg shorter than the arrowhead just looks like a kink, so level it out.
      double levelTolerance = ArrowSize * scale;

      double centerX = requestedCenter.X;
      double centerY = requestedCenter.Y;
      double z = tip.Z;
      double toTipX = tip.X - centerX;
      double toTipY = tip.Y - centerY;

      if (Math.Abs(toTipX) <= halfWidth)
      {
        // Tip is within the hexagon's width: stack the note straight above or below it.
        centerX = tip.X;
        Point3d center = new Point3d(centerX, centerY, z);
        if (Math.Abs(toTipY) <= halfHeight)
        {
          return new KeyNoteLeaderLayout(center, new Point3d[0]);
        }

        double edgeY = centerY + (toTipY > 0.0 ? halfHeight : -halfHeight);
        return new KeyNoteLeaderLayout(center, new[] { tip, new Point3d(centerX, edgeY, z) });
      }

      double exitX = centerX + (toTipX > 0.0 ? halfWidth : -halfWidth);
      if (Math.Abs(toTipY) < levelTolerance)
      {
        centerY = tip.Y;
        return new KeyNoteLeaderLayout(
          new Point3d(centerX, centerY, z),
          new[] { tip, new Point3d(exitX, centerY, z) });
      }

      return new KeyNoteLeaderLayout(
        new Point3d(centerX, centerY, z),
        new[] { tip, new Point3d(tip.X, centerY, z), new Point3d(exitX, centerY, z) });
    }
  }

  // Drags the keyed note around a fixed leader tip, showing the leader as it will be drawn.
  internal sealed class KeyNoteLeaderJig : DrawJig, IDisposable
  {
    private readonly Database _db;
    private readonly BlockReference _note;
    private readonly Point3d _tip;
    private readonly double _scale;
    private Point3d _requestedCenter;
    private KeyNoteLeaderLayout _layout;
    private bool _hasSample;
    private bool _pendingRedraw;

    // Takes ownership of the transient note and disposes it.
    internal KeyNoteLeaderJig(
      Database db,
      BlockReference note,
      Point3d tip,
      double scale,
      KeyNoteLeaderEnd end)
    {
      _db = db;
      _note = note;
      _tip = tip;
      _scale = scale;
      End = end;
      _layout = KeyNoteLeaderLayout.Compute(tip, tip, scale);
    }

    internal KeyNoteLeaderLayout Layout => _layout;

    internal KeyNoteLeaderEnd End { get; private set; }

    internal static KeyNoteLeaderEnd ParseEnd(string keyword, KeyNoteLeaderEnd current)
    {
      if (string.Equals(keyword, "Arrow", StringComparison.OrdinalIgnoreCase))
      {
        return KeyNoteLeaderEnd.Arrow;
      }
      if (string.Equals(keyword, "Circle", StringComparison.OrdinalIgnoreCase))
      {
        return KeyNoteLeaderEnd.Circle;
      }
      return current;
    }

    internal void ApplyKeyword(string keyword)
    {
      End = ParseEnd(keyword, End);
      _pendingRedraw = true;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions
      {
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted
      };
      options.SetMessageAndKeywords(
        $"\nSpecify keyed note location or [Arrow/Circle] (current: {End}): ",
        "Arrow Circle");

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Keyword)
      {
        return SamplerStatus.OK;
      }
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }
      if (_hasSample && !_pendingRedraw && result.Value.DistanceTo(_requestedCenter) < 1e-6)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _pendingRedraw = false;
      _requestedCenter = result.Value;
      _layout = KeyNoteLeaderLayout.Compute(_tip, _requestedCenter, _scale);
      _note.Position = _layout.NoteCenter;
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (!_hasSample) return true;

      draw.Geometry.Draw(_note);
      foreach (Entity preview in BuildLeaderPreview())
      {
        try
        {
          draw.Geometry.Draw(preview);
        }
        finally
        {
          preview.Dispose();
        }
      }
      return true;
    }

    // Plain lines and a circle rather than a Leader, so the preview does not
    // depend on the arrowhead block or the dimension style.
    private List<Entity> BuildLeaderPreview()
    {
      List<Entity> preview = new List<Entity>();
      if (!_layout.HasLeader) return preview;

      Point3d[] path = _layout.Path;
      for (int i = 0; i < path.Length - 1; i++)
      {
        preview.Add(CreatePreviewLine(path[i], path[i + 1]));
      }

      double size = KeyNoteLeaderLayout.ArrowSize * _scale;
      if (End == KeyNoteLeaderEnd.Circle)
      {
        Circle circle = new Circle(path[0], Vector3d.ZAxis, size / 2.0);
        circle.SetDatabaseDefaults(_db);
        circle.Layer = _note.Layer;
        circle.ColorIndex = 256;
        preview.Add(circle);
        return preview;
      }

      Vector3d direction = (path[0] - path[1]).GetNormal();
      Vector3d perpendicular = Vector3d.ZAxis.CrossProduct(direction);
      Point3d baseCenter = path[0] - direction * size;
      // AutoCAD's closed arrowhead is a third as wide as it is long.
      Point3d left = baseCenter + perpendicular * (size / 6.0);
      Point3d right = baseCenter - perpendicular * (size / 6.0);
      preview.Add(CreatePreviewLine(path[0], left));
      preview.Add(CreatePreviewLine(path[0], right));
      preview.Add(CreatePreviewLine(left, right));
      return preview;
    }

    private Line CreatePreviewLine(Point3d start, Point3d end)
    {
      Line line = new Line(start, end);
      line.SetDatabaseDefaults(_db);
      line.Layer = _note.Layer;
      line.ColorIndex = 256;
      return line;
    }

    public void Dispose()
    {
      _note.Dispose();
    }
  }
}
