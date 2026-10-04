using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
  internal enum SymbolSide
  {
    North,
    East,
    South,
    West
  }

  // Previews a symbol at the cursor. Symbol geometry is built around the origin
  // in plotted inches, the same units as the inserted block definition.
  internal sealed class SymbolLocationJig : DrawJig, IDisposable
  {
    private readonly SymbolPreview _preview;
    private readonly string _message;
    private readonly string _globalKeywords;
    private readonly double _rotation;
    private readonly double _scale;
    private Point3d _location = Point3d.Origin;
    private bool _hasSample;

    public SymbolLocationJig(
      List<Entity> previewEntities,
      string message,
      string globalKeywords,
      double scale,
      double rotation
    )
    {
      _preview = new SymbolPreview(previewEntities);
      _message = message;
      _globalKeywords = globalKeywords;
      _rotation = rotation;
      _scale = scale;
    }

    public Point3d Location => _location;

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions
      {
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted,
      };
      if (string.IsNullOrEmpty(_globalKeywords))
      {
        options.Message = _message;
      }
      else
      {
        options.SetMessageAndKeywords(_message, _globalKeywords);
      }

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Keyword)
      {
        return SamplerStatus.OK;
      }
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }
      if (_hasSample && result.Value.DistanceTo(_location) < 1e-7)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _location = result.Value;
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (_hasSample)
      {
        _preview.Draw(
          draw,
          SymbolGeometry.CreateInsertTransform(_location, _rotation, _scale)
        );
      }
      return true;
    }

    public void Dispose()
    {
      _preview.Dispose();
    }
  }

  // Lets the user pick a side of a placed symbol with the cursor or N/E/S/W keywords.
  internal sealed class SymbolSideJig : DrawJig, IDisposable
  {
    private readonly Dictionary<SymbolSide, SymbolPreview> _previews =
      new Dictionary<SymbolSide, SymbolPreview>();
    private readonly Point3d _location;
    private readonly Matrix3d _transform;
    private readonly double _rotation;
    private readonly double _minimumPickDistance;
    private readonly string _message;
    private SymbolSide _side;

    public SymbolSideJig(
      Func<SymbolSide, List<Entity>> createPreviewEntities,
      SymbolSide initialSide,
      string message,
      Point3d location,
      double scale,
      double rotation,
      double minimumPickDistance
    )
    {
      foreach (SymbolSide side in Enum.GetValues(typeof(SymbolSide)))
      {
        _previews[side] = new SymbolPreview(createPreviewEntities(side));
      }
      _side = initialSide;
      _location = location;
      _rotation = rotation;
      _transform = SymbolGeometry.CreateInsertTransform(location, rotation, scale);
      // Close to the insertion point the direction is too twitchy to read, so keep the current side.
      _minimumPickDistance = minimumPickDistance;
      _message = message;
    }

    public SymbolSide Side => _side;

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions
      {
        BasePoint = _location,
        UseBasePoint = true,
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NullResponseAccepted
          | UserInputControls.NoNegativeResponseAccepted,
        Cursor = CursorType.RubberBand,
      };
      options.SetMessageAndKeywords(_message, "North East South West");

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status == PromptStatus.Keyword || result.Status == PromptStatus.None)
      {
        return SamplerStatus.OK;
      }
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }

      SymbolSide side = ResolveSide(result.Value);
      if (side == _side)
      {
        return SamplerStatus.NoChange;
      }

      _side = side;
      return SamplerStatus.OK;
    }

    private SymbolSide ResolveSide(Point3d cursor)
    {
      // Measure in the symbol's own frame so North is always toward the top of the UCS.
      Vector3d offset = (cursor - _location).RotateBy(-_rotation, Vector3d.ZAxis);
      if (Math.Sqrt(offset.X * offset.X + offset.Y * offset.Y) < _minimumPickDistance)
      {
        return _side;
      }
      if (Math.Abs(offset.X) > Math.Abs(offset.Y))
      {
        return offset.X > 0.0 ? SymbolSide.East : SymbolSide.West;
      }
      return offset.Y > 0.0 ? SymbolSide.North : SymbolSide.South;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      _previews[_side].Draw(draw, _transform);
      return true;
    }

    public void Dispose()
    {
      foreach (SymbolPreview preview in _previews.Values)
      {
        preview.Dispose();
      }
    }
  }

  internal sealed class SymbolPreview : IDisposable
  {
    private readonly List<Entity> _entities;

    internal SymbolPreview(List<Entity> entities)
    {
      _entities = entities;
    }

    internal void Draw(WorldDraw draw, Matrix3d transform)
    {
      draw.Geometry.PushModelTransform(transform);
      try
      {
        foreach (Entity entity in _entities)
        {
          draw.Geometry.Draw(entity);
        }
      }
      finally
      {
        draw.Geometry.PopModelTransform();
      }
    }

    public void Dispose()
    {
      foreach (Entity entity in _entities)
      {
        entity.Dispose();
      }
    }
  }

  internal static class SymbolGeometry
  {
    internal const short RedColorIndex = 1;
    internal const short YellowColorIndex = 2;
    internal const short CyanColorIndex = 4;
    internal const short WhiteColorIndex = 7;

    // Symbol sizes are specified in model inches at 1/4" = 1'-0"; dividing by
    // this converts them to the plotted inches the block definitions use.
    internal const double QuarterScaleModelInchesPerPlottedInch = 48.0;

    internal static Matrix3d CreateInsertTransform(
      Point3d location,
      double rotation,
      double scale
    )
    {
      return Matrix3d.Displacement(location.GetAsVector())
        * Matrix3d.Rotation(rotation, Vector3d.ZAxis, Point3d.Origin)
        * Matrix3d.Scaling(scale, Point3d.Origin);
    }

    // Rotation that turns geometry drawn facing north toward the given side.
    internal static double ResolveSideRotation(SymbolSide side)
    {
      switch (side)
      {
        case SymbolSide.East:
          return -Math.PI / 2.0;
        case SymbolSide.South:
          return Math.PI;
        case SymbolSide.West:
          return Math.PI / 2.0;
        default:
          return 0.0;
      }
    }

    internal static Entity Prepare(Database database, Entity entity, short colorIndex)
    {
      entity.SetDatabaseDefaults(database);
      entity.LayerId = database.LayerZero;
      entity.LinetypeId = database.ByLayerLinetype;
      entity.LineWeight = LineWeight.ByLayer;
      entity.ColorIndex = colorIndex;
      return entity;
    }

    // Uses the drawing's existing definition when present so office-edited
    // symbols are respected; otherwise builds it from the supplied geometry.
    internal static ObjectId EnsureBlock(
      Transaction transaction,
      Database database,
      string blockName,
      Func<List<Entity>> createEntities
    )
    {
      BlockTable blockTable = (BlockTable)transaction.GetObject(
        database.BlockTableId,
        OpenMode.ForRead);
      if (blockTable.Has(blockName))
      {
        return blockTable[blockName];
      }

      blockTable.UpgradeOpen();
      BlockTableRecord blockDefinition = new BlockTableRecord
      {
        Name = blockName,
        Origin = Point3d.Origin,
      };
      ObjectId blockDefinitionId = blockTable.Add(blockDefinition);
      transaction.AddNewlyCreatedDBObject(blockDefinition, true);

      foreach (Entity entity in createEntities())
      {
        blockDefinition.AppendEntity(entity);
        transaction.AddNewlyCreatedDBObject(entity, true);
      }

      return blockDefinitionId;
    }

    // Keeps symbols upright and N/E/S/W aligned with the current UCS.
    internal static double ResolveUcsRotation(Editor editor)
    {
      Vector3d ucsXAxis = Vector3d.XAxis.TransformBy(editor.CurrentUserCoordinateSystem);
      return Math.Atan2(ucsXAxis.Y, ucsXAxis.X);
    }
  }
}
