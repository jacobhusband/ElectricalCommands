using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private static string _lastEquipmentTagValue;

    [CommandMethod("EQUIPMENTTAG", CommandFlags.Modal)]
    [CommandMethod("ETAG", CommandFlags.Modal)]
    public void EquipmentTagCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!TryResolveKnPlacementScale(ed, db, "ETAG", out double scaleDenom)) return;

      string tagValue = PromptEquipmentTagValue(ed, _lastEquipmentTagValue);
      if (string.IsNullOrEmpty(tagValue))
      {
        ed.WriteMessage("\nETAG canceled.");
        return;
      }

      PlaceEquipmentTag(db, ed, "ETAG", scaleDenom, tagValue, false);
    }

    // Asks for the tagged point and the tag location, then draws the tag. With
    // allowEdit, the tagged-point prompt shows the tag text and offers an Edit
    // keyword, so a misread tag can be corrected without starting over.
    private static void PlaceEquipmentTag(
      Database db,
      Editor ed,
      string commandName,
      double scaleDenom,
      string tagValue,
      bool allowEdit)
    {
      Point3d tip;
      while (true)
      {
        PromptPointOptions tipOptions = new PromptPointOptions(string.Empty)
        {
          AllowNone = false
        };
        if (allowEdit)
        {
          tipOptions.SetMessageAndKeywords(
            $"\nSpecify point to tag for {tagValue} or [Edit]: ",
            "Edit");
        }
        else
        {
          tipOptions.Message = "\nSpecify point to tag: ";
        }

        PromptPointResult tipResult = ed.GetPoint(tipOptions);
        if (tipResult.Status == PromptStatus.Keyword)
        {
          tagValue = PromptEquipmentTagValue(ed, tagValue);
          if (string.IsNullOrEmpty(tagValue))
          {
            ed.WriteMessage($"\n{commandName} canceled.");
            return;
          }
          continue;
        }
        if (tipResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage($"\n{commandName} canceled.");
          return;
        }
        tip = tipResult.Value;
        break;
      }

      EquipmentTagLayout layout;
      EquipmentTagFactory factory;
      try
      {
        factory = new EquipmentTagFactory(db, EnsureKnTextStyle(db), tagValue, scaleDenom);
        EquipmentTagJig jig = new EquipmentTagJig(factory, tip);
        if (ed.Drag(jig).Status != PromptStatus.OK)
        {
          ed.WriteMessage($"\n{commandName} canceled.");
          return;
        }
        layout = jig.Layout;
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\n{commandName} jig error: {ex.Message}");
        return;
      }

      try
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
          foreach (Entity entity in factory.CreateEntities(layout))
          {
            space.AppendEntity(entity);
            tr.AddNewlyCreatedDBObject(entity, true);
          }
          tr.Commit();
        }

        _lastEquipmentTagValue = tagValue;
        ed.WriteMessage(
          layout.HasLine
            ? $"\nPlaced equipment tag {tagValue} with a line to the tagged point."
            : $"\nPlaced equipment tag {tagValue}; the oval covers the tagged point, so no line was drawn.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\n{commandName} placement error: {ex.Message}");
      }
    }

    private static string PromptEquipmentTagValue(Editor ed, string defaultValue)
    {
      string prompt = string.IsNullOrEmpty(defaultValue)
        ? "\nEnter equipment tag (e.g., EF-1, AHU-2): "
        : $"\nEnter equipment tag (e.g., EF-1, AHU-2) <{defaultValue}>: ";

      PromptResult pr = ed.GetString(new PromptStringOptions(prompt)
      {
        AllowSpaces = true
      });
      if (pr.Status != PromptStatus.OK) return null;

      string value = pr.StringResult?.Trim() ?? string.Empty;
      return value.Length == 0 ? defaultValue : value;
    }

    // Builds the tag's MText, oval, and line on the current layer, sized like a
    // KEYED_NOTE: plotted-inch dimensions times the viewport scale.
    internal sealed class EquipmentTagFactory
    {
      private readonly Database _db;
      private readonly ObjectId _textStyleId;
      private readonly string _contents;
      private readonly double _textHeight;

      internal EquipmentTagFactory(Database db, ObjectId textStyleId, string tagValue, double scale)
      {
        _db = db;
        _textStyleId = textStyleId;
        _contents = EscapeMTextPlainText(tagValue);
        _textHeight = KnAttributeHeight * scale;

        // The smallest ellipse around a w x h box has semi-axes sqrt(2) * w/2 and
        // sqrt(2) * h/2. Pad the text first so it does not touch the oval, and
        // never let the oval get narrower than it is tall.
        using (MText sample = CreateText(Point3d.Origin))
        {
          double padding = _textHeight * 0.35;
          double halfWidth = Math.Max(sample.ActualWidth, _textHeight) / 2.0 + padding;
          double halfHeight = Math.Max(sample.ActualHeight, _textHeight) / 2.0 + padding;
          MinorRadius = Math.Sqrt(2.0) * halfHeight;
          MajorRadius = Math.Max(Math.Sqrt(2.0) * halfWidth, MinorRadius);
        }
      }

      internal double MajorRadius { get; }
      internal double MinorRadius { get; }

      // Caller owns the returned entities: append them to the database or dispose them.
      internal List<Entity> CreateEntities(EquipmentTagLayout layout)
      {
        List<Entity> entities = new List<Entity> { CreateText(layout.Center) };

        Ellipse oval = new Ellipse(
          layout.Center,
          Vector3d.ZAxis,
          Vector3d.XAxis * MajorRadius,
          MinorRadius / MajorRadius,
          0.0,
          2.0 * Math.PI);
        oval.SetDatabaseDefaults(_db);
        oval.ColorIndex = 256;
        entities.Add(oval);

        if (layout.HasLine)
        {
          Line line = new Line(layout.LineStart, layout.Tip);
          line.SetDatabaseDefaults(_db);
          line.ColorIndex = 256;
          entities.Add(line);
        }
        return entities;
      }

      private MText CreateText(Point3d center)
      {
        MText text = new MText();
        text.SetDatabaseDefaults(_db);
        text.Contents = _contents;
        text.TextHeight = _textHeight;
        text.Width = 0.0;
        text.Rotation = 0.0;
        text.Attachment = AttachmentPoint.MiddleCenter;
        text.Location = center;
        text.ColorIndex = 256;
        if (!_textStyleId.IsNull)
        {
          text.TextStyleId = _textStyleId;
        }
        return text;
      }
    }
  }

  // Where the equipment tag sits and where its line leaves the oval. The line
  // aims at the oval's center, so it starts where that ray crosses the oval.
  internal readonly struct EquipmentTagLayout
  {
    private EquipmentTagLayout(Point3d center, Point3d tip, Point3d lineStart, bool hasLine)
    {
      Center = center;
      Tip = tip;
      LineStart = lineStart;
      HasLine = hasLine;
    }

    internal Point3d Center { get; }
    internal Point3d Tip { get; }
    internal Point3d LineStart { get; }

    // False when the oval covers the tagged point.
    internal bool HasLine { get; }

    internal static EquipmentTagLayout Compute(
      Point3d tip,
      Point3d requestedCenter,
      double majorRadius,
      double minorRadius)
    {
      Point3d center = new Point3d(requestedCenter.X, requestedCenter.Y, tip.Z);
      Vector3d toTip = tip - center;

      // 1 on the oval, less than 1 inside it.
      double reach = Math.Sqrt(
        (toTip.X * toTip.X) / (majorRadius * majorRadius) +
        (toTip.Y * toTip.Y) / (minorRadius * minorRadius));
      if (reach <= 1.0 + 1e-9)
      {
        return new EquipmentTagLayout(center, tip, tip, false);
      }

      return new EquipmentTagLayout(center, tip, center + toTip / reach, true);
    }
  }

  // Drags the equipment tag around a fixed tagged point, showing the oval and line as they will be drawn.
  internal sealed class EquipmentTagJig : DrawJig
  {
    private readonly GeneralCommands.EquipmentTagFactory _factory;
    private readonly Point3d _tip;
    private EquipmentTagLayout _layout;
    private bool _hasSample;

    internal EquipmentTagJig(GeneralCommands.EquipmentTagFactory factory, Point3d tip)
    {
      _factory = factory;
      _tip = tip;
      _layout = EquipmentTagLayout.Compute(tip, tip, factory.MajorRadius, factory.MinorRadius);
    }

    internal EquipmentTagLayout Layout => _layout;

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions options = new JigPromptPointOptions("\nSpecify tag location: ")
      {
        UserInputControls = UserInputControls.Accept3dCoordinates
          | UserInputControls.NoNegativeResponseAccepted
      };

      PromptPointResult result = prompts.AcquirePoint(options);
      if (result.Status != PromptStatus.OK)
      {
        return SamplerStatus.Cancel;
      }
      if (_hasSample && result.Value.DistanceTo(_layout.Center) < 1e-6)
      {
        return SamplerStatus.NoChange;
      }

      _hasSample = true;
      _layout = EquipmentTagLayout.Compute(
        _tip, result.Value, _factory.MajorRadius, _factory.MinorRadius);
      return SamplerStatus.OK;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (!_hasSample) return true;

      foreach (Entity preview in _factory.CreateEntities(_layout))
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
  }
}
