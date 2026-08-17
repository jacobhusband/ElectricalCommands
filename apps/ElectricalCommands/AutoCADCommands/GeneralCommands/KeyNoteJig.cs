using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace ElectricalCommands
{
  internal sealed class KeyNoteJig : EntityJig
  {
    public enum Anchor { Center, Left, Right, Top, Bottom }

    private const double HexHalfWidth = 0.1252;
    private const double HexHalfHeight = 0.108426;

    private Point3d _pickPoint = Point3d.Origin;
    private bool _firstSample = true;
    private bool _pendingRedraw = false;
    private Anchor _anchor = Anchor.Center;
    private readonly double _scale;

    public KeyNoteJig(BlockReference blockReference, double scale) : base(blockReference)
    {
      _scale = scale;
    }

    public Point3d InsertionPoint => ComputePosition();

    public Anchor CurrentAnchor => _anchor;

    public bool ApplyKeyword(string keyword)
    {
      if (string.IsNullOrEmpty(keyword)) return false;
      switch (keyword)
      {
        case "Center": _anchor = Anchor.Center; break;
        case "Left": _anchor = Anchor.Left; break;
        case "Right": _anchor = Anchor.Right; break;
        case "Up": _anchor = Anchor.Top; break;
        case "Down": _anchor = Anchor.Bottom; break;
        default: return false;
      }
      _pendingRedraw = true;
      return true;
    }

    private Vector3d GetAnchorOffset()
    {
      switch (_anchor)
      {
        case Anchor.Left: return new Vector3d(-HexHalfWidth, 0.0, 0.0);
        case Anchor.Right: return new Vector3d(HexHalfWidth, 0.0, 0.0);
        case Anchor.Top: return new Vector3d(0.0, HexHalfHeight, 0.0);
        case Anchor.Bottom: return new Vector3d(0.0, -HexHalfHeight, 0.0);
        default: return new Vector3d(0.0, 0.0, 0.0);
      }
    }

    private Point3d ComputePosition()
    {
      Vector3d offset = GetAnchorOffset() * _scale;
      return _pickPoint - offset;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions opts = new JigPromptPointOptions(
        $"\nSpecify keyed note insertion point (anchor: {_anchor})")
      {
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted
      };
      opts.Keywords.Add("Center");
      opts.Keywords.Add("Left");
      opts.Keywords.Add("Right");
      opts.Keywords.Add("Up");
      opts.Keywords.Add("Down");

      PromptPointResult res = prompts.AcquirePoint(opts);

      if (res.Status == PromptStatus.Keyword)
      {
        return SamplerStatus.OK;
      }
      if (res.Status == PromptStatus.Cancel || res.Status == PromptStatus.Error)
      {
        return SamplerStatus.Cancel;
      }
      if (res.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }

      if (!_firstSample && !_pendingRedraw && res.Value.DistanceTo(_pickPoint) < 1e-6)
      {
        return SamplerStatus.NoChange;
      }

      _firstSample = false;
      _pendingRedraw = false;
      _pickPoint = res.Value;
      return SamplerStatus.OK;
    }

    protected override bool Update()
    {
      ((BlockReference)Entity).Position = ComputePosition();
      return true;
    }
  }
}
