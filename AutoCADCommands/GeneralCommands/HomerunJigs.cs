using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using System;

namespace ElectricalCommands
{
  internal sealed class HomerunArrowTipJig : DrawJig, IDisposable
  {
    private readonly Point3d _panelPoint;
    private readonly HomerunArrowPreview _arrow;
    private Point3d _tipPoint = Point3d.Origin;
    private bool _hasSample;

    public HomerunArrowTipJig(
      Database database,
      Point3d panelPoint,
      double arrowSize
    )
    {
      _panelPoint = panelPoint;
      _arrow = new HomerunArrowPreview(
        database,
        _tipPoint,
        _panelPoint - _tipPoint,
        arrowSize
      );
    }

    public Point3d TipPoint => _tipPoint;

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions(
        "\nSpecify the arrow tip (preview points toward the panel): "
      )
      {
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted,
      };

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Cancel || result.Status == PromptStatus.Error)
      {
        return SamplerStatus.Cancel;
      }
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }
      if (_hasSample && result.Value.DistanceTo(_tipPoint) < 1e-7)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _tipPoint = result.Value;
      _arrow.Update(_tipPoint, _panelPoint - _tipPoint);
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      _arrow.Draw(draw);
      return true;
    }

    public void Dispose()
    {
      _arrow.Dispose();
    }
  }

  internal sealed class HomerunMiddleJig : DrawJig, IDisposable
  {
    private readonly Point3d _tipPoint;
    private readonly Line _headSegment;
    private readonly HomerunArrowPreview _arrow;
    private Point3d _middlePoint;
    private bool _hasSample;

    public HomerunMiddleJig(
      Database database,
      Point3d tipPoint,
      Point3d panelPoint,
      double arrowSize
    )
    {
      _tipPoint = tipPoint;
      Vector3d towardPanel = panelPoint - tipPoint;
      Vector3d initialDirection = HomerunArrowPreview.NormalizeOrFallback(
        towardPanel,
        Vector3d.XAxis
      );
      _middlePoint = tipPoint - (initialDirection * Math.Max(arrowSize * 2.0, 1.0));
      _headSegment = HomerunEntityFactory.CreatePreviewLine(
        database,
        _tipPoint,
        _middlePoint
      );
      _arrow = new HomerunArrowPreview(
        database,
        _tipPoint,
        _tipPoint - _middlePoint,
        arrowSize
      );
    }

    public Point3d MiddlePoint => _middlePoint;

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions(
        "\nSpecify the middle point of the arrow: "
      )
      {
        BasePoint = _tipPoint,
        UseBasePoint = true,
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted,
        Cursor = CursorType.RubberBand,
      };

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Cancel || result.Status == PromptStatus.Error)
      {
        return SamplerStatus.Cancel;
      }
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }
      if (_hasSample && result.Value.DistanceTo(_middlePoint) < 1e-7)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _middlePoint = result.Value;
      _headSegment.EndPoint = _middlePoint;
      _arrow.Update(_tipPoint, _tipPoint - _middlePoint);
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      draw.Geometry.Draw(_headSegment);
      _arrow.Draw(draw);
      return true;
    }

    public void Dispose()
    {
      _headSegment.Dispose();
      _arrow.Dispose();
    }
  }

  internal sealed class HomerunBaseJig : DrawJig, IDisposable
  {
    private readonly Point3d _middlePoint;
    private readonly Line _headSegment;
    private readonly Line _tailSegment;
    private readonly HomerunArrowPreview _arrow;
    private Point3d _basePoint;
    private bool _hasSample;

    public HomerunBaseJig(
      Database database,
      Point3d tipPoint,
      Point3d middlePoint,
      double arrowSize
    )
    {
      _middlePoint = middlePoint;
      Vector3d continuation = middlePoint - tipPoint;
      _basePoint = middlePoint + continuation;
      _headSegment = HomerunEntityFactory.CreatePreviewLine(
        database,
        tipPoint,
        middlePoint
      );
      _tailSegment = HomerunEntityFactory.CreatePreviewLine(
        database,
        middlePoint,
        _basePoint
      );
      _arrow = new HomerunArrowPreview(
        database,
        tipPoint,
        tipPoint - middlePoint,
        arrowSize
      );
    }

    public Point3d BasePoint => _basePoint;

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions(
        "\nSpecify the base point of the arrow: "
      )
      {
        BasePoint = _middlePoint,
        UseBasePoint = true,
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted,
        Cursor = CursorType.RubberBand,
      };

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Cancel || result.Status == PromptStatus.Error)
      {
        return SamplerStatus.Cancel;
      }
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }
      if (_hasSample && result.Value.DistanceTo(_basePoint) < 1e-7)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _basePoint = result.Value;
      _tailSegment.EndPoint = _basePoint;
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      draw.Geometry.Draw(_headSegment);
      draw.Geometry.Draw(_tailSegment);
      _arrow.Draw(draw);
      return true;
    }

    public void Dispose()
    {
      _headSegment.Dispose();
      _tailSegment.Dispose();
      _arrow.Dispose();
    }
  }

  internal sealed class HomerunTextJig : DrawJig, IDisposable
  {
    private readonly Line _headSegment;
    private readonly Line _tailSegment;
    private readonly HomerunArrowPreview _arrow;
    private readonly MText _text;
    private Point3d _textPoint;
    private bool _hasSample;

    public HomerunTextJig(
      Database database,
      Point3d headPoint,
      Point3d centerPoint,
      Point3d tailPoint,
      Point3d initialTextPoint,
      double symbolSize,
      string contents,
      ObjectId previewTextStyleId
    )
    {
      _textPoint = initialTextPoint;
      _headSegment = HomerunEntityFactory.CreatePreviewLine(database, headPoint, centerPoint);
      _tailSegment = HomerunEntityFactory.CreatePreviewLine(database, centerPoint, tailPoint);
      _arrow = new HomerunArrowPreview(
        database,
        headPoint,
        headPoint - centerPoint,
        symbolSize
      );
      _text = HomerunEntityFactory.CreateMText(
        database,
        initialTextPoint,
        symbolSize,
        contents,
        previewTextStyleId
      );
    }

    public Point3d TextPoint => _textPoint;

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions(
        "\nSpecify the panel label location: "
      )
      {
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted,
        Cursor = CursorType.RubberBand,
      };

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Cancel || result.Status == PromptStatus.Error)
      {
        return SamplerStatus.Cancel;
      }
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }
      if (_hasSample && result.Value.DistanceTo(_textPoint) < 1e-7)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _textPoint = result.Value;
      _text.Location = _textPoint;
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      draw.Geometry.Draw(_headSegment);
      draw.Geometry.Draw(_tailSegment);
      _arrow.Draw(draw);
      draw.Geometry.Draw(_text);
      return true;
    }

    public void Dispose()
    {
      _headSegment.Dispose();
      _tailSegment.Dispose();
      _arrow.Dispose();
      _text.Dispose();
    }

  }

  internal sealed class HomerunArrowPreview : IDisposable
  {
    private readonly Line _leftWing;
    private readonly Line _rightWing;
    private readonly Line _base;
    private readonly double _arrowSize;

    internal HomerunArrowPreview(
      Database database,
      Point3d tipPoint,
      Vector3d pointingDirection,
      double arrowSize
    )
    {
      _arrowSize = Math.Max(Math.Abs(arrowSize), 1e-6);
      _leftWing = HomerunEntityFactory.CreatePreviewLine(database, tipPoint, tipPoint);
      _rightWing = HomerunEntityFactory.CreatePreviewLine(database, tipPoint, tipPoint);
      _base = HomerunEntityFactory.CreatePreviewLine(database, tipPoint, tipPoint);
      Update(tipPoint, pointingDirection);
    }

    internal void Update(Point3d tipPoint, Vector3d pointingDirection)
    {
      Vector3d direction = NormalizeOrFallback(pointingDirection, Vector3d.XAxis);
      Vector3d perpendicular = Vector3d.ZAxis.CrossProduct(direction);
      if (perpendicular.Length < 1e-9)
      {
        perpendicular = Vector3d.YAxis.CrossProduct(direction);
      }
      perpendicular = NormalizeOrFallback(perpendicular, Vector3d.YAxis);

      Point3d baseCenter = tipPoint - (direction * _arrowSize);
      Vector3d halfWidth = perpendicular * (_arrowSize * 0.35);
      Point3d leftPoint = baseCenter + halfWidth;
      Point3d rightPoint = baseCenter - halfWidth;

      _leftWing.StartPoint = tipPoint;
      _leftWing.EndPoint = leftPoint;
      _rightWing.StartPoint = tipPoint;
      _rightWing.EndPoint = rightPoint;
      _base.StartPoint = leftPoint;
      _base.EndPoint = rightPoint;
    }

    internal void Draw(WorldDraw draw)
    {
      draw.Geometry.Draw(_leftWing);
      draw.Geometry.Draw(_rightWing);
      draw.Geometry.Draw(_base);
    }

    internal static Vector3d NormalizeOrFallback(Vector3d vector, Vector3d fallback)
    {
      return vector.Length < 1e-9 ? fallback.GetNormal() : vector.GetNormal();
    }

    public void Dispose()
    {
      _leftWing.Dispose();
      _rightWing.Dispose();
      _base.Dispose();
    }
  }

  internal static class HomerunEntityFactory
  {
    private const short CyanColorIndex = 4;

    public static Line CreatePreviewLine(Database database, Point3d startPoint, Point3d endPoint)
    {
      Line line = new Line(startPoint, endPoint);
      line.SetDatabaseDefaults(database);
      line.ColorIndex = CyanColorIndex;
      return line;
    }

    public static Leader CreateLeader(
      Database database,
      Point3d headPoint,
      Point3d centerPoint,
      Point3d tailPoint,
      double arrowSize
    )
    {
      Leader leader = new Leader();
      leader.SetDatabaseDefaults(database);
      leader.ColorIndex = CyanColorIndex;
      leader.HasArrowHead = true;
      leader.IsSplined = true;
      leader.Dimasz = arrowSize;
      leader.Dimldrblk = ObjectId.Null;
      leader.AppendVertex(headPoint);
      leader.AppendVertex(centerPoint);
      leader.AppendVertex(tailPoint);
      return leader;
    }

    public static MText CreateMText(
      Database database,
      Point3d location,
      double textHeight,
      string contents,
      ObjectId textStyleId
    )
    {
      MText text = new MText();
      text.SetDatabaseDefaults(database);
      text.ColorIndex = CyanColorIndex;
      text.Location = location;
      text.Contents = contents;
      text.Attachment = AttachmentPoint.MiddleCenter;
      text.TextHeight = textHeight;
      text.Width = 0.0;
      // A new MText is already NoColumns. AutoCAD throws eNotApplicable when
      // ColumnType.NoColumns is explicitly assigned before the entity is database-resident.
      if (!textStyleId.IsNull)
      {
        text.TextStyleId = textStyleId;
      }
      return text;
    }
  }
}
