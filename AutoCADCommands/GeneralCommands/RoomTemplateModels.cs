using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;

namespace ElectricalCommands
{
  internal enum TemplateMemberKind
  {
    Receptacle,
    Symbol,
    Note,
    Leader
  }

  internal sealed class TemplatePoint
  {
    public double X { get; set; }
    public double Y { get; set; }
  }

  // One captured object. Receptacles and symbols sit at a drawing-unit offset from
  // the group's base point, so real-world spacing survives a change of drawing
  // scale. Notes and leaders hang off a parent symbol at an offset in plotted
  // inches, so they keep their paper size at any scale.
  internal sealed class TemplateMember
  {
    public TemplateMemberKind Kind { get; set; }
    public string BlockName { get; set; } = string.Empty;
    public string Layer { get; set; } = string.Empty;
    public double X { get; set; }
    public double Y { get; set; }

    // Receptacles and symbols: rotation relative to the capture frame (UCS-aligned).
    // Notes: rotation relative to the UCS at capture.
    public double RelativeRotation { get; set; }

    // Block scale divided by the drawing's block scale at capture.
    public double ScaleRatio { get; set; } = 1.0;
    // Older templates stored X alone. Their Y/Z were normally positive, even
    // when X was negative after MIRROR; never turn that into a three-axis flip.
    public double? ScaleRatioY { get; set; }
    public double? ScaleRatioZ { get; set; }
    public List<TemplateAttribute> AttributeGeometry { get; set; } = new List<TemplateAttribute>();

    public Dictionary<string, string> Attributes { get; set; } =
      new Dictionary<string, string>();
    public Dictionary<string, string> DynamicProperties { get; set; } =
      new Dictionary<string, string>();
    public string VisibilityState { get; set; } = string.Empty;

    // Notes and leaders: the symbol they move with (-1 = the group's base point).
    public int ParentIndex { get; set; } = -1;

    // Leaders: the note at the far end of the leader (-1 = none) and the arrow style.
    public int NoteIndex { get; set; } = -1;
    public bool CircleEnd { get; set; }

    // Leaders: the captured path (tip first) as offsets from the parent in plotted
    // inches. Only used when the leader has no note to re-route toward.
    public List<TemplatePoint> LeaderPath { get; set; } = new List<TemplatePoint>();

    // Receptacles: which template circuit the receptacle belongs to (-1 = none).
    public int CircuitSlot { get; set; } = -1;

    public bool IsSymbol =>
      Kind == TemplateMemberKind.Receptacle || Kind == TemplateMemberKind.Symbol;
  }

  internal sealed class TemplateGroup
  {
    public string Name { get; set; } = string.Empty;

    // Direction the group faces (into the room) relative to the capture frame,
    // in radians. Null when the group has no facing, such as a ceiling device.
    public double? FacingAngle { get; set; }

    public List<TemplateMember> Members { get; set; } = new List<TemplateMember>();
  }

  // Attribute geometry in the block's local coordinates. Values alone cannot
  // reproduce attributes moved with grips or corrected by MIRRTEXT.
  internal sealed class TemplateAttribute
  {
    public string Tag { get; set; }
    public double[] Position { get; set; }
    public double[] AlignmentPoint { get; set; }
    public double Rotation { get; set; }
    public double Height { get; set; }
    public double WidthFactor { get; set; }
    public double Oblique { get; set; }
    public int HorizontalMode { get; set; }
    public int VerticalMode { get; set; }
    public bool MirroredX { get; set; }
    public bool MirroredY { get; set; }
  }

  internal sealed class RoomTemplate
  {
    public string Name { get; set; } = string.Empty;
    public double CaptureBlockScale { get; set; } = 1.0;

    // One entry per circuit; informational (the label text the circuit had at capture).
    public List<string> CircuitLabels { get; set; } = new List<string>();

    public List<TemplateGroup> Groups { get; set; } = new List<TemplateGroup>();
  }

  // A member after the template has been placed in the drawing's coordinates.
  internal sealed class ResolvedTemplateMember
  {
    internal TemplateMember Member { get; set; }
    internal Point3d Position { get; set; }
    internal double Rotation { get; set; }
    internal Scale3d ScaleFactors { get; set; }

    // Leaders: the vertices to draw, tip first. Empty when the note covers the tip.
    internal Point3d[] LeaderPath { get; set; } = new Point3d[0];
  }

  internal static class RoomTemplateGeometry
  {
    // Converts a world point into the UCS-aligned frame used by captured groups.
    internal static Vector3d ToLocal(Point3d point, Point3d origin, double ucsRotation)
    {
      Vector3d offset = point - origin;
      return new Vector3d(offset.X, offset.Y, 0.0).RotateBy(-ucsRotation, Vector3d.ZAxis);
    }

    // Rotation to add to a group so that its facing direction points along
    // wallNormal. Groups without a facing line their X axis up with the wall.
    internal static double RotationForWall(
      TemplateGroup group,
      double wallAngle,
      double wallNormal,
      double ucsRotation
    )
    {
      return group.FacingAngle.HasValue
        ? NormalizeAngle(wallNormal - group.FacingAngle.Value - ucsRotation)
        : NormalizeAngle(wallAngle - ucsRotation);
    }

    internal static double NormalizeAngle(double angle)
    {
      double twoPi = 2.0 * Math.PI;
      angle %= twoPi;
      if (angle > Math.PI)
      {
        angle -= twoPi;
      }
      else if (angle <= -Math.PI)
      {
        angle += twoPi;
      }
      return angle;
    }

    // Lays a group out in the drawing. userRotation turns the whole group about its
    // base point; ucsRotation keeps it aligned with the current UCS.
    internal static List<ResolvedTemplateMember> Resolve(
      TemplateGroup group,
      Point3d origin,
      double ucsRotation,
      double userRotation,
      double blockScale
    )
    {
      double frame = ucsRotation + userRotation;
      List<ResolvedTemplateMember> resolved = new List<ResolvedTemplateMember>();
      foreach (TemplateMember member in group.Members)
      {
        resolved.Add(new ResolvedTemplateMember { Member = member });
      }

      // Symbols first: everything else is positioned relative to them.
      for (int index = 0; index < group.Members.Count; index++)
      {
        TemplateMember member = group.Members[index];
        if (!member.IsSymbol)
        {
          continue;
        }

        Vector3d offset = new Vector3d(member.X, member.Y, 0.0).RotateBy(frame, Vector3d.ZAxis);
        resolved[index].Position = origin + offset;
        resolved[index].Rotation = frame + member.RelativeRotation;
        resolved[index].ScaleFactors = ResolveScale(member, blockScale);
      }

      for (int index = 0; index < group.Members.Count; index++)
      {
        TemplateMember member = group.Members[index];
        if (member.Kind != TemplateMemberKind.Note && member.Kind != TemplateMemberKind.Leader)
        {
          continue;
        }

        resolved[index].Position = AnnotationPoint(
          group,
          resolved,
          member.ParentIndex,
          member.X,
          member.Y,
          origin,
          frame,
          blockScale
        );
        resolved[index].Rotation = ucsRotation + member.RelativeRotation;
        resolved[index].ScaleFactors = ResolveScale(member, blockScale);
      }

      // Keep the captured path and note center. Recomputing a KNL layout here
      // changes the saved group even when it is placed without any rotation.
      for (int index = 0; index < group.Members.Count; index++)
      {
        TemplateMember member = group.Members[index];
        if (member.Kind != TemplateMemberKind.Leader)
        {
          continue;
        }

        List<Point3d> path = new List<Point3d>();
        foreach (TemplatePoint vertex in member.LeaderPath)
        {
          path.Add(
            AnnotationPoint(
              group,
              resolved,
              member.ParentIndex,
              vertex.X,
              vertex.Y,
              origin,
              frame,
              blockScale
            )
          );
        }
        resolved[index].LeaderPath = path.ToArray();
      }

      return resolved;
    }

    private static Scale3d ResolveScale(TemplateMember member, double blockScale)
    {
      // Legacy notes lost their MIRRTEXT-corrected attributes. The hexagon is
      // symmetric, so a positive scale restores readable text without moving it.
      double x = member.Kind == TemplateMemberKind.Note && !member.ScaleRatioY.HasValue
        ? Math.Abs(member.ScaleRatio) : member.ScaleRatio;
      return new Scale3d(blockScale * x,
        blockScale * (member.ScaleRatioY ?? Math.Abs(member.ScaleRatio)),
        blockScale * (member.ScaleRatioZ ?? Math.Abs(member.ScaleRatio)));
    }

    private static Point3d AnnotationPoint(
      TemplateGroup group,
      List<ResolvedTemplateMember> resolved,
      int parentIndex,
      double plottedX,
      double plottedY,
      Point3d origin,
      double frame,
      double blockScale
    )
    {
      Point3d anchor =
        parentIndex >= 0 && parentIndex < group.Members.Count && group.Members[parentIndex].IsSymbol
          ? resolved[parentIndex].Position
          : origin;
      Vector3d offset = new Vector3d(plottedX * blockScale, plottedY * blockScale, 0.0).RotateBy(
        frame,
        Vector3d.ZAxis
      );
      return anchor + offset;
    }
  }
}
