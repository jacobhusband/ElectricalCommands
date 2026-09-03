using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Newtonsoft.Json;

namespace ElectricalCommands
{
  public enum SwitchType
  {
    Standard,
    Dimmer,
    Occupancy
  }

  public enum SwitchOrientation
  {
    North,
    East,
    South,
    West
  }

  public sealed class StoredVector3d
  {
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }

    public StoredVector3d() { }

    public StoredVector3d(double x, double y, double z)
    {
      X = x;
      Y = y;
      Z = z;
    }

    public Vector3d ToVector3d() => new Vector3d(X, Y, Z);

    public static StoredVector3d FromVector3d(Vector3d v) =>
      new StoredVector3d(v.X, v.Y, v.Z);

    public static StoredVector3d FromPoint3d(Point3d p) =>
      new StoredVector3d(p.X, p.Y, p.Z);
  }

  public sealed class SwitchTextDefinition
  {
    public string TextString { get; set; } = string.Empty;
    public string TextStyleName { get; set; } = "Standard";
    public double Height { get; set; } = 0.09375;
    public double Rotation { get; set; } = 0.0;
    public string Layer { get; set; } = "0";
    public int ColorIndex { get; set; } = 256; // ByLayer
    public bool IsMText { get; set; }
    public double Width { get; set; }
    public int AttachmentPoint { get; set; } = (int)Autodesk.AutoCAD.DatabaseServices.AttachmentPoint.BaseLeft;
    public int HorizontalMode { get; set; } = (int)TextHorizontalMode.TextLeft;
    public int VerticalMode { get; set; } = (int)TextVerticalMode.TextBase;
    public StoredVector3d RelativeOffset { get; set; } = new StoredVector3d();

    public SwitchTextDefinition Clone()
    {
      return new SwitchTextDefinition
      {
        TextString = TextString,
        TextStyleName = TextStyleName,
        Height = Height,
        Rotation = Rotation,
        Layer = Layer,
        ColorIndex = ColorIndex,
        IsMText = IsMText,
        Width = Width,
        AttachmentPoint = AttachmentPoint,
        HorizontalMode = HorizontalMode,
        VerticalMode = VerticalMode,
        RelativeOffset = new StoredVector3d(RelativeOffset.X, RelativeOffset.Y, RelativeOffset.Z)
      };
    }
  }

  public sealed class SwitchAttributeDefinition
  {
    public string Tag { get; set; } = string.Empty;
    public string TextString { get; set; } = string.Empty;
    public StoredVector3d RelativeOffset { get; set; } = new StoredVector3d();
  }

  public sealed class SwitchBlockDefinition
  {
    public string BlockName { get; set; } = string.Empty;
    public double Rotation { get; set; }
    public double ScaleX { get; set; } = 1.0;
    public double ScaleY { get; set; } = 1.0;
    public double ScaleZ { get; set; } = 1.0;
    public string Layer { get; set; } = "0";
    public string VisibilityState { get; set; } = string.Empty;
    public List<SwitchAttributeDefinition> Attributes { get; set; } =
      new List<SwitchAttributeDefinition>();

    public SwitchBlockDefinition Clone()
    {
      var clone = new SwitchBlockDefinition
      {
        BlockName = BlockName,
        Rotation = Rotation,
        ScaleX = ScaleX,
        ScaleY = ScaleY,
        ScaleZ = ScaleZ,
        Layer = Layer,
        VisibilityState = VisibilityState,
        Attributes = new List<SwitchAttributeDefinition>()
      };
      if (Attributes != null)
      {
        foreach (var attr in Attributes)
        {
          clone.Attributes.Add(new SwitchAttributeDefinition
          {
            Tag = attr.Tag,
            TextString = attr.TextString,
            RelativeOffset = new StoredVector3d(attr.RelativeOffset.X, attr.RelativeOffset.Y, attr.RelativeOffset.Z)
          });
        }
      }
      return clone;
    }
  }

  public sealed class SwitchOrientationConfig
  {
    public SwitchBlockDefinition Block { get; set; }
    public List<SwitchTextDefinition> TextObjects { get; set; } =
      new List<SwitchTextDefinition>();

    [JsonIgnore]
    public bool IsConfigured =>
      Block != null && !string.IsNullOrWhiteSpace(Block.BlockName);

    public SwitchOrientationConfig Clone()
    {
      var clone = new SwitchOrientationConfig
      {
        Block = Block?.Clone(),
        TextObjects = new List<SwitchTextDefinition>()
      };
      if (TextObjects != null)
      {
        foreach (var txt in TextObjects)
        {
          clone.TextObjects.Add(txt.Clone());
        }
      }
      return clone;
    }
  }

  public sealed class SwitchTypeConfig
  {
    public string TypeName { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public SwitchOrientationConfig North { get; set; } = new SwitchOrientationConfig();
    public SwitchOrientationConfig East { get; set; } = new SwitchOrientationConfig();
    public SwitchOrientationConfig South { get; set; } = new SwitchOrientationConfig();
    public SwitchOrientationConfig West { get; set; } = new SwitchOrientationConfig();

    public SwitchOrientationConfig GetOrientation(SwitchOrientation orientation)
    {
      switch (orientation)
      {
        case SwitchOrientation.North: return North ?? (North = new SwitchOrientationConfig());
        case SwitchOrientation.East: return East ?? (East = new SwitchOrientationConfig());
        case SwitchOrientation.South: return South ?? (South = new SwitchOrientationConfig());
        case SwitchOrientation.West: return West ?? (West = new SwitchOrientationConfig());
        default: return North ?? (North = new SwitchOrientationConfig());
      }
    }

    public void SetOrientation(SwitchOrientation orientation, SwitchOrientationConfig config)
    {
      switch (orientation)
      {
        case SwitchOrientation.North: North = config; break;
        case SwitchOrientation.East: East = config; break;
        case SwitchOrientation.South: South = config; break;
        case SwitchOrientation.West: West = config; break;
      }
    }

    [JsonIgnore]
    public bool HasAnyConfiguration =>
      (North != null && North.IsConfigured) ||
      (East != null && East.IsConfigured) ||
      (South != null && South.IsConfigured) ||
      (West != null && West.IsConfigured);

    public SwitchTypeConfig Clone()
    {
      return new SwitchTypeConfig
      {
        TypeName = TypeName,
        DisplayName = DisplayName,
        North = North?.Clone() ?? new SwitchOrientationConfig(),
        East = East?.Clone() ?? new SwitchOrientationConfig(),
        South = South?.Clone() ?? new SwitchOrientationConfig(),
        West = West?.Clone() ?? new SwitchOrientationConfig()
      };
    }
  }

  public sealed class ProjectSwitchSettings
  {
    public int Version { get; set; } = 1;
    public string UpdatedUtc { get; set; } = string.Empty;
    public string SourceDrawing { get; set; } = string.Empty;
    public Dictionary<string, SwitchTypeConfig> Switches { get; set; } =
      new Dictionary<string, SwitchTypeConfig>(StringComparer.OrdinalIgnoreCase);

    public ProjectSwitchSettings()
    {
      EnsureDefaults();
    }

    public void EnsureDefaults()
    {
      if (!Switches.ContainsKey(SwitchType.Standard.ToString()))
      {
        Switches[SwitchType.Standard.ToString()] = new SwitchTypeConfig
        {
          TypeName = SwitchType.Standard.ToString(),
          DisplayName = "Standard Switch"
        };
      }
      if (!Switches.ContainsKey(SwitchType.Dimmer.ToString()))
      {
        Switches[SwitchType.Dimmer.ToString()] = new SwitchTypeConfig
        {
          TypeName = SwitchType.Dimmer.ToString(),
          DisplayName = "Dimmer Switch"
        };
      }
      if (!Switches.ContainsKey(SwitchType.Occupancy.ToString()))
      {
        Switches[SwitchType.Occupancy.ToString()] = new SwitchTypeConfig
        {
          TypeName = SwitchType.Occupancy.ToString(),
          DisplayName = "Occupancy Switch"
        };
      }
    }

    public SwitchTypeConfig GetTypeConfig(SwitchType type)
    {
      EnsureDefaults();
      string key = type.ToString();
      if (Switches.TryGetValue(key, out var config))
      {
        return config;
      }
      config = new SwitchTypeConfig
      {
        TypeName = key,
        DisplayName = type == SwitchType.Standard ? "Standard Switch" :
                      type == SwitchType.Dimmer ? "Dimmer Switch" : "Occupancy Switch"
      };
      Switches[key] = config;
      return config;
    }
  }
}
