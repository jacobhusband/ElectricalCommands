using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace ElectricalCommands
{
  internal enum TemplateJigInput
  {
    None,
    Point,
    Keyword,
    Enter
  }

  // Drags a whole template group at the cursor. Space, Enter and right-click come
  // back as an Enter so the command can turn the group a quarter turn.
  internal sealed class RoomTemplateGroupJig : DrawJig, IDisposable
  {
    private const string Keywords = "Rotate Flip Wall Angle Skip Back Done";
    private const short NoteColorIndex = 2;

    private readonly TemplateGroup _group;
    private readonly double _blockScale;
    private readonly double _ucsRotation;
    private readonly Database _database;
    private readonly Dictionary<int, BlockReference> _previewBlocks =
      new Dictionary<int, BlockReference>();
    private Point3d _location = Point3d.Origin;
    private double _userRotation;
    private bool _hasSample;
    private bool _pendingRedraw;

    // createPreviewBlock returns an unattached block reference for a symbol or note
    // (or null when its definition is unavailable); the jig owns and disposes it.
    internal RoomTemplateGroupJig(
      Database database,
      TemplateGroup group,
      Func<TemplateMember, BlockReference> createPreviewBlock,
      double blockScale,
      double ucsRotation,
      double userRotation
    )
    {
      _database = database;
      _group = group;
      _blockScale = blockScale;
      _ucsRotation = ucsRotation;
      _userRotation = userRotation;

      for (int index = 0; index < group.Members.Count; index++)
      {
        TemplateMember member = group.Members[index];
        if (member.Kind == TemplateMemberKind.Leader)
        {
          continue;
        }

        BlockReference preview = createPreviewBlock(member);
        if (preview != null)
        {
          _previewBlocks[index] = preview;
        }
      }
    }

    internal string Message { get; set; } = "\nSpecify group location: ";

    internal Point3d Location => _location;

    internal bool HasLocation => _hasSample;

    internal TemplateJigInput LastInput { get; private set; }

    internal string LastKeyword { get; private set; } = string.Empty;

    internal double UserRotation
    {
      get => _userRotation;
      set
      {
        _userRotation = value;
        _pendingRedraw = true;
      }
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions
      {
        UserInputControls =
          UserInputControls.Accept3dCoordinates
          | UserInputControls.NullResponseAccepted
          | UserInputControls.NoNegativeResponseAccepted,
      };
      options.SetMessageAndKeywords(Message, Keywords);

      PromptPointResult result = prompts.AcquirePoint(options);
      LastKeyword = string.Empty;
      if (result.Status == PromptStatus.Keyword)
      {
        LastInput = TemplateJigInput.Keyword;
        LastKeyword = result.StringResult ?? string.Empty;
        return SamplerStatus.OK;
      }
      if (result.Status == PromptStatus.None)
      {
        LastInput = TemplateJigInput.Enter;
        return SamplerStatus.OK;
      }
      if (result.Status != PromptStatus.OK)
      {
        LastInput = TemplateJigInput.None;
        return SamplerStatus.Cancel;
      }

      LastInput = TemplateJigInput.Point;
      if (_hasSample && !_pendingRedraw && result.Value.DistanceTo(_location) < 1e-7)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _pendingRedraw = false;
      _location = result.Value;
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (!_hasSample)
      {
        return true;
      }

      List<ResolvedTemplateMember> resolved = RoomTemplateGeometry.Resolve(
        _group,
        _location,
        _ucsRotation,
        _userRotation,
        _blockScale
      );
      for (int index = 0; index < resolved.Count; index++)
      {
        ResolvedTemplateMember member = resolved[index];
        if (member.Member.Kind == TemplateMemberKind.Leader)
        {
          DrawLeader(draw, member);
          continue;
        }

        if (_previewBlocks.TryGetValue(index, out BlockReference block))
        {
          block.Position = member.Position;
          block.Rotation = member.Rotation;
          block.ScaleFactors = new Scale3d(member.Scale);
          draw.Geometry.Draw(block);
        }
      }
      return true;
    }

    // Plain lines and a circle rather than a Leader, so the preview does not depend on
    // the arrowhead block or the dimension style (the same approach KNL's jig takes).
    private void DrawLeader(WorldDraw draw, ResolvedTemplateMember leader)
    {
      Point3d[] path = leader.LeaderPath;
      if (path == null || path.Length < 2)
      {
        return;
      }

      List<Entity> pieces = new List<Entity>();
      for (int index = 0; index < path.Length - 1; index++)
      {
        pieces.Add(CreateLine(path[index], path[index + 1]));
      }

      double size = KeyNoteLeaderLayout.ArrowSize * _blockScale;
      if (leader.Member.CircleEnd)
      {
        Circle circle = new Circle(path[0], Vector3d.ZAxis, size / 2.0);
        circle.ColorIndex = NoteColorIndex;
        pieces.Add(circle);
      }
      else
      {
        Vector3d direction = (path[0] - path[1]).GetNormal();
        Vector3d perpendicular = Vector3d.ZAxis.CrossProduct(direction);
        Point3d baseCenter = path[0] - direction * size;
        // AutoCAD's closed arrowhead is a third as wide as it is long.
        Point3d left = baseCenter + perpendicular * (size / 6.0);
        Point3d right = baseCenter - perpendicular * (size / 6.0);
        pieces.Add(CreateLine(path[0], left));
        pieces.Add(CreateLine(path[0], right));
        pieces.Add(CreateLine(left, right));
      }

      foreach (Entity piece in pieces)
      {
        try
        {
          draw.Geometry.Draw(piece);
        }
        finally
        {
          piece.Dispose();
        }
      }
    }

    private static Line CreateLine(Point3d start, Point3d end)
    {
      return new Line(start, end) { ColorIndex = NoteColorIndex };
    }

    public void Dispose()
    {
      foreach (BlockReference block in _previewBlocks.Values)
      {
        block.Dispose();
      }
      _previewBlocks.Clear();
    }
  }
}
