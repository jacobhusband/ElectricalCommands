using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private static SymbolSide _lastTxtlSide = SymbolSide.East;
    private static KeyNoteLeaderEnd _lastTxtlEnd = KeyNoteLeaderEnd.Arrow;

    [CommandMethod("TXTL", CommandFlags.Modal)]
    public static void InsertScaledTextLeader()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage("\nTXTL requires a drawing scale. Run SETSCALE (SS) first.");
        return;
      }

      PromptResult textResult = ed.GetString(new PromptStringOptions("\nEnter text: ")
      {
        AllowSpaces = true,
      });
      if (textResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nTXTL canceled.");
        return;
      }
      string text = (textResult.StringResult ?? string.Empty).Trim();
      if (text.Length == 0)
      {
        ed.WriteMessage("\nText cannot be blank.");
        return;
      }

      KeyNoteLeaderEnd end = _lastTxtlEnd;
      Point3d tip;
      while (true)
      {
        PromptPointOptions options = new PromptPointOptions(string.Empty);
        options.SetMessageAndKeywords(
          $"\nSpecify leader tip or [Arrow/Circle] (current: {end}): ",
          "Arrow Circle");
        PromptPointResult result = ed.GetPoint(options);
        if (result.Status == PromptStatus.Keyword)
        {
          end = KeyNoteLeaderJig.ParseEnd(result.StringResult, end);
          continue;
        }
        if (result.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nTXTL canceled.");
          return;
        }
        tip = result.Value;
        break;
      }

      try
      {
        string contents = EscapeMTextPlainText(text);
        double textHeight = ResolveHomerunSymbolSize(scale.PaperInchesPerModelFoot);
        double leaderScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
        double ucsRotation = SymbolGeometry.ResolveUcsRotation(ed);
        ObjectId textStyleId = EnsureHomerunTextStyle(db);
        Func<SymbolSide, MText> createText = side =>
          CreateScaledText(db, contents, side, textHeight, textStyleId);

        using (ScaledTextLeaderJig jig = new ScaledTextLeaderJig(
          db, createText, tip, ucsRotation, leaderScale, _lastTxtlSide, end))
        {
          while (true)
          {
            PromptResult result = ed.Drag(jig);
            if (result.Status == PromptStatus.Keyword)
            {
              jig.ApplyKeyword(result.StringResult);
              continue;
            }
            if (result.Status != PromptStatus.OK)
            {
              ed.WriteMessage("\nTXTL canceled.");
              return;
            }
            if (jig.Layout.Path.Length < 2)
            {
              ed.WriteMessage("\nPlace the text farther from the leader tip.");
              continue;
            }
            break;
          }

          using (Transaction tr = db.TransactionManager.StartTransaction())
          {
            BlockTableRecord space = (BlockTableRecord)tr.GetObject(
              db.CurrentSpaceId, OpenMode.ForWrite);
            MText mtext = createText(jig.Side);
            mtext.TransformBy(SymbolGeometry.CreateInsertTransform(
              jig.Layout.Location, ucsRotation, 1.0));
            space.AppendEntity(mtext);
            tr.AddNewlyCreatedDBObject(mtext, true);

            Leader leader = CreateAnnotationLeader(db, jig.Layout.Path, jig.End, leaderScale);
            leader.LayerId = mtext.LayerId;
            space.AppendEntity(leader);
            tr.AddNewlyCreatedDBObject(leader, true);
            tr.Commit();
          }

          _lastTxtlSide = jig.Side;
          _lastTxtlEnd = jig.End;
        }
        ed.WriteMessage($"\nPlaced scaled MText with a leader at {scale.DisplayText}.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to place TXTL: {ex.Message}");
      }
    }
  }

  internal readonly struct ScaledTextLeaderLayout
  {
    internal ScaledTextLeaderLayout(Point3d location, Point3d[] path)
    {
      Location = location;
      Path = path;
    }

    internal Point3d Location { get; }
    internal Point3d[] Path { get; }

    // Work along the text's baseline so vertical text and a rotated UCS use
    // the same edge/gap rules. Choose the nearer edge to avoid crossing the text.
    internal static ScaledTextLeaderLayout Compute(
      Point3d tip, Point3d requestedLocation, SymbolSide side,
      double ucsRotation, double width, double height)
    {
      bool vertical = side == SymbolSide.North || side == SymbolSide.South;
      bool backward = side == SymbolSide.West || side == SymbolSide.South;
      double rotation = ucsRotation + (vertical ? Math.PI / 2.0 : 0.0);
      Vector3d along = new Vector3d(Math.Cos(rotation), Math.Sin(rotation), 0);
      Vector3d across = Vector3d.ZAxis.CrossProduct(along);
      Point3d location = new Point3d(requestedLocation.X, requestedLocation.Y, tip.Z);
      Vector3d offset = tip - location;
      double x = offset.DotProduct(along);
      double y = offset.DotProduct(across);
      double minX = backward ? -width : 0;
      double maxX = backward ? 0 : width;
      double gap = height * 0.5;
      Point3d edge;
      Point3d elbow;

      if (x >= minX - gap && x <= maxX + gap)
      {
        // Tip is aligned with the text's width: leave its upper/lower edge.
        if (Math.Abs(y) <= height / 2.0 + gap)
          return new ScaledTextLeaderLayout(location, Array.Empty<Point3d>());
        double edgeX = Math.Max(minX, Math.Min(maxX, x));
        edge = location + along * edgeX + across * Math.Sign(y) * (height / 2.0 + gap);
        elbow = tip;
      }
      else
      {
        if (Math.Abs(y) < height)
        {
          // Like KNL, level out a leg that is shorter than the arrowhead.
          location += across * y;
        }
        edge = location + along * (x < minX ? minX - gap : maxX + gap);
        elbow = location + along * x;
      }

      Point3d[] path = tip.DistanceTo(elbow) < 1e-7
        ? new[] { tip, edge }
        : new[] { tip, elbow, edge };
      if (tip.DistanceTo(path[1]) < 1e-7)
        return new ScaledTextLeaderLayout(location, Array.Empty<Point3d>());
      return new ScaledTextLeaderLayout(location, path);
    }
  }

  internal sealed class ScaledTextLeaderJig : DrawJig, IDisposable
  {
    private readonly Database _db;
    private readonly Func<SymbolSide, MText> _createText;
    private readonly Point3d _tip;
    private readonly double _rotation;
    private readonly double _leaderScale;
    private MText _text;
    private Point3d _requestedLocation;
    private bool _hasSample;
    private bool _pendingRedraw;

    internal ScaledTextLeaderJig(Database db, Func<SymbolSide, MText> createText,
      Point3d tip, double rotation, double leaderScale, SymbolSide side, KeyNoteLeaderEnd end)
    {
      _db = db;
      _createText = createText;
      _tip = tip;
      _rotation = rotation;
      _leaderScale = leaderScale;
      _requestedLocation = tip;
      Side = side;
      End = end;
      UpdatePreview();
    }

    internal SymbolSide Side { get; private set; }
    internal KeyNoteLeaderEnd End { get; private set; }
    internal ScaledTextLeaderLayout Layout { get; private set; }

    internal void ApplyKeyword(string keyword)
    {
      if (Enum.TryParse(keyword, true, out SymbolSide side)) Side = side;
      End = KeyNoteLeaderJig.ParseEnd(keyword, End);
      _pendingRedraw = true;
      UpdatePreview();
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions
      {
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NullResponseAccepted
          | UserInputControls.NoNegativeResponseAccepted,
      };
      options.SetMessageAndKeywords(
        $"\nSpecify text location or [North/East/South/West/Arrow/Circle] (current: {Side}, {End}): ",
        "North East South West Arrow Circle");
      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Keyword) return SamplerStatus.OK;
      if (result.Status != PromptStatus.OK) return SamplerStatus.Cancel;
      if (_hasSample && !_pendingRedraw && result.Value.DistanceTo(_requestedLocation) < 1e-7)
        return SamplerStatus.NoChange;
      _hasSample = true;
      _pendingRedraw = false;
      _requestedLocation = result.Value;
      UpdatePreview();
      return SamplerStatus.OK;
    }

    private void UpdatePreview()
    {
      _text?.Dispose();
      _text = _createText(Side);
      Layout = ScaledTextLeaderLayout.Compute(_tip, _requestedLocation, Side,
        _rotation, _text.ActualWidth, _text.ActualHeight);
      _text.TransformBy(SymbolGeometry.CreateInsertTransform(Layout.Location, _rotation, 1.0));
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (!_hasSample) return true;
      draw.Geometry.Draw(_text);
      Point3d[] path = Layout.Path;
      for (int i = 0; i < path.Length - 1; i++) DrawLine(draw, path[i], path[i + 1]);
      if (path.Length < 2) return true;

      double size = KeyNoteLeaderLayout.ArrowSize * _leaderScale;
      if (End == KeyNoteLeaderEnd.Circle)
      {
        using (Circle circle = new Circle(path[0], Vector3d.ZAxis, size / 2.0))
        {
          circle.SetDatabaseDefaults(_db);
          circle.LayerId = _text.LayerId;
          circle.ColorIndex = 256;
          draw.Geometry.Draw(circle);
        }
      }
      else
      {
        Vector3d direction = (path[0] - path[1]).GetNormal();
        Vector3d perpendicular = Vector3d.ZAxis.CrossProduct(direction);
        Point3d center = path[0] - direction * size;
        Point3d left = center + perpendicular * (size / 6.0);
        Point3d right = center - perpendicular * (size / 6.0);
        DrawLine(draw, path[0], left);
        DrawLine(draw, path[0], right);
        DrawLine(draw, left, right);
      }
      return true;
    }

    private void DrawLine(WorldDraw draw, Point3d start, Point3d end)
    {
      using (Line line = new Line(start, end))
      {
        line.SetDatabaseDefaults(_db);
        line.LayerId = _text.LayerId;
        line.ColorIndex = 256;
        draw.Geometry.Draw(line);
      }
    }

    public void Dispose() => _text?.Dispose();
  }
}
