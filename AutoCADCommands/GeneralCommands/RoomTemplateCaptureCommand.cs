using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    // A label further than this from a receptacle (plotted inches) is not its circuit label.
    private const double TemplateLabelCutoffPlottedInches = 2.0;

    // A keyed note this close (plotted inches) to the end of a leader is that leader's note.
    private const double TemplateNoteMatchPlottedInches = 0.5;

    // Remembered for the AutoCAD session so RTP offers the template just captured.
    private static string _lastRoomTemplateName = string.Empty;

    // A leader's geometry, read while its transaction was open.
    private sealed class CapturedLeader
    {
      internal string Layer { get; set; } = string.Empty;
      internal bool CircleEnd { get; set; }
      internal List<Point3d> Vertices { get; } = new List<Point3d>();
    }

    private sealed class CapturedReceptacle
    {
      internal ObjectId ObjectId { get; set; }
      internal Point3d Position { get; set; }
      internal TemplateMember Member { get; set; }
    }

    [CommandMethod("RTCAPTURE", CommandFlags.Modal)]
    [CommandMethod("RTC", CommandFlags.Modal)]
    public static void CaptureRoomTemplate()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!TryResolveTemplateScale(db, ed, "RTC", out var scale)) return;
      double blockScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
      double ucsRotation = SymbolGeometry.ResolveUcsRotation(ed);

      if (!TryPromptRoomTemplateName(db, ed, out string templateName)) return;

      RoomTemplate template = new RoomTemplate
      {
        Name = templateName,
        CaptureBlockScale = blockScale,
      };
      List<CapturedReceptacle> receptacles = new List<CapturedReceptacle>();
      HashSet<ObjectId> claimed = new HashSet<ObjectId>();

      while (true)
      {
        int groupNumber = template.Groups.Count + 1;
        string finishHint = template.Groups.Count > 0 ? " (Enter when there are no more groups)" : string.Empty;
        PromptSelectionResult selection = ed.GetSelection(
          new PromptSelectionOptions
          {
            MessageForAdding =
              $"\nSelect the symbols, keyed notes and leaders in group {groupNumber}{finishHint}: ",
          },
          new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "INSERT,LEADER") })
        );
        if (selection.Status != PromptStatus.OK || selection.Value == null)
        {
          if (template.Groups.Count == 0)
          {
            ed.WriteMessage("\nRTC canceled.");
            return;
          }
          break;
        }

        List<ObjectId> selected = new List<ObjectId>();
        foreach (ObjectId id in selection.Value.GetObjectIds())
        {
          if (claimed.Add(id))
          {
            selected.Add(id);
          }
        }
        if (selected.Count == 0)
        {
          ed.WriteMessage("\nEvery selected object already belongs to another group.");
          continue;
        }

        PromptPointOptions baseOptions = new PromptPointOptions(
          "\nSpecify the group base point (on the wall, or the room center): "
        );
        PromptPointResult baseResult = ed.GetPoint(baseOptions);
        if (baseResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nRTC canceled.");
          return;
        }

        PromptPointOptions facingOptions = new PromptPointOptions(
          "\nSpecify the direction the group faces into the room, or Enter for none: "
        )
        {
          BasePoint = baseResult.Value,
          UseBasePoint = true,
          AllowNone = true,
        };
        PromptPointResult facingResult = ed.GetPoint(facingOptions);
        if (facingResult.Status != PromptStatus.OK && facingResult.Status != PromptStatus.None)
        {
          ed.WriteMessage("\nRTC canceled.");
          return;
        }

        string defaultName = $"Group {groupNumber}";
        PromptStringOptions nameOptions = new PromptStringOptions(
          $"\nName for this group <{defaultName}>: "
        )
        {
          AllowSpaces = true,
        };
        PromptResult nameResult = ed.GetString(nameOptions);
        if (nameResult.Status != PromptStatus.OK && nameResult.Status != PromptStatus.None)
        {
          ed.WriteMessage("\nRTC canceled.");
          return;
        }
        string groupName = (nameResult.StringResult ?? string.Empty).Trim();
        if (groupName.Length == 0)
        {
          groupName = defaultName;
        }

        Matrix3d ucsToWcs = ed.CurrentUserCoordinateSystem;
        Point3d origin = baseResult.Value.TransformBy(ucsToWcs);
        double? facingAngle = null;
        if (facingResult.Status == PromptStatus.OK)
        {
          Vector3d facing = (facingResult.Value - baseResult.Value).TransformBy(ucsToWcs);
          if (facing.Length > 1e-6)
          {
            facingAngle = RoomTemplateGeometry.NormalizeAngle(
              Math.Atan2(facing.Y, facing.X) - ucsRotation
            );
          }
        }

        try
        {
          TemplateGroup group = BuildTemplateGroup(
            db,
            selected,
            groupName,
            origin,
            facingAngle,
            ucsRotation,
            blockScale,
            receptacles
          );
          if (group.Members.Count == 0)
          {
            ed.WriteMessage("\nNothing in that selection could be captured; the group was skipped.");
            continue;
          }
          template.Groups.Add(group);
          ed.WriteMessage($"\nCaptured {DescribeTemplateGroup(group)}.");
        }
        catch (System.Exception ex)
        {
          ed.WriteMessage($"\nUnable to capture the group: {ex.Message}");
          return;
        }
      }

      if (receptacles.Count > 0 && !TryAssignTemplateCircuits(db, ed, template, receptacles))
      {
        ed.WriteMessage("\nRTC canceled.");
        return;
      }

      try
      {
        RoomTemplateStore.Save(db, template);
        _lastRoomTemplateName = template.Name;
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to save the template: {ex.Message}");
        return;
      }

      ed.WriteMessage(
        $"\nSaved room template {template.Name}: {template.Groups.Count} group(s), "
          + $"{receptacles.Count} receptacle(s), {template.CircuitLabels.Count} circuit(s). "
          + "Run RTP in each room to place it."
      );
    }

    // Shared by RTC and RTP: the same scale rules R uses, including taking the scale
    // from the viewport being edited.
    private static bool TryResolveTemplateScale(
      Database db,
      Editor ed,
      string commandName,
      out ElectricalDrawingSettingsStore.ScaleSetting scale
    )
    {
      if (
        TrySetReceptacleScaleFromActiveViewport(
          db,
          ed,
          out scale,
          out bool isEditingViewport,
          out string viewportScaleError
        )
      )
      {
        ed.WriteMessage(
          $"\nDrawing scale automatically set to {scale.DisplayText} from the active viewport."
        );
        DraftingPalette.Refresh();
        DraftingPalette.SetStatus($"Scale set to {scale.DisplayText} from the active viewport.");
        return true;
      }
      if (isEditingViewport)
      {
        ed.WriteMessage(
          $"\n{commandName} could not determine the active viewport scale: {viewportScaleError}"
        );
        return false;
      }
      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out scale))
      {
        ed.WriteMessage($"\n{commandName} requires a drawing scale. Run SETSCALE (SS) first.");
        return false;
      }
      return true;
    }

    // Names are limited to letters, digits, hyphens and underscores so they can be typed
    // at a prompt and used as drawing dictionary keys.
    private static string NormalizeTemplateName(string name)
    {
      StringBuilder normalized = new StringBuilder();
      foreach (char letter in (name ?? string.Empty).Trim().ToUpperInvariant())
      {
        normalized.Append(char.IsLetterOrDigit(letter) || letter == '-' ? letter : '_');
      }
      return normalized.ToString();
    }

    private static bool TryPromptRoomTemplateName(Database db, Editor ed, out string templateName)
    {
      templateName = string.Empty;
      List<string> existing = RoomTemplateStore.ListNames(db);
      string suggestion = $"ROOM{existing.Count + 1}";
      string existingText = existing.Count > 0 ? $" Existing: {string.Join(", ", existing)}." : string.Empty;
      PromptStringOptions options = new PromptStringOptions(
        $"\nEnter a name for the room template.{existingText} <{suggestion}>: "
      )
      {
        AllowSpaces = true,
      };
      PromptResult result = ed.GetString(options);
      if (result.Status != PromptStatus.OK && result.Status != PromptStatus.None)
      {
        ed.WriteMessage("\nRTC canceled.");
        return false;
      }

      templateName = NormalizeTemplateName(result.StringResult);
      if (templateName.Length == 0)
      {
        templateName = suggestion;
      }

      if (RoomTemplateStore.TryLoad(db, templateName, out _))
      {
        PromptKeywordOptions overwrite = new PromptKeywordOptions(
          $"\nTemplate {templateName} already exists. Overwrite it? [Yes/No] <No>: ",
          "Yes No"
        )
        {
          AllowNone = true,
        };
        overwrite.Keywords.Default = "No";
        PromptResult overwriteResult = ed.GetKeywords(overwrite);
        if (overwriteResult.Status != PromptStatus.OK || overwriteResult.StringResult != "Yes")
        {
          ed.WriteMessage("\nRTC canceled; the existing template was not changed.");
          return false;
        }
      }
      return true;
    }

    private static string DescribeTemplateGroup(TemplateGroup group)
    {
      int symbols = 0;
      int notes = 0;
      foreach (TemplateMember member in group.Members)
      {
        if (member.IsSymbol)
        {
          symbols++;
        }
        else if (member.Kind == TemplateMemberKind.Note)
        {
          notes++;
        }
      }
      return $"\"{group.Name}\": {symbols} symbol(s), {notes} keyed note(s)"
        + (group.FacingAngle.HasValue ? ", with a facing direction" : string.Empty);
    }

    private static TemplateGroup BuildTemplateGroup(
      Database db,
      List<ObjectId> objectIds,
      string groupName,
      Point3d origin,
      double? facingAngle,
      double ucsRotation,
      double blockScale,
      List<CapturedReceptacle> receptacles
    )
    {
      TemplateGroup group = new TemplateGroup { Name = groupName, FacingAngle = facingAngle };
      List<TemplateMember> symbols = new List<TemplateMember>();
      List<Point3d> symbolPositions = new List<Point3d>();
      List<CapturedReceptacle> groupReceptacles = new List<CapturedReceptacle>();
      List<TemplateMember> notes = new List<TemplateMember>();
      List<Point3d> notePositions = new List<Point3d>();
      List<CapturedLeader> leaders = new List<CapturedLeader>();

      using (Transaction transaction = db.TransactionManager.StartTransaction())
      {
        foreach (ObjectId id in objectIds)
        {
          Entity entity = transaction.GetObject(id, OpenMode.ForRead) as Entity;
          if (entity is Leader sourceLeader)
          {
            CapturedLeader captured = new CapturedLeader
            {
              Layer = sourceLeader.Layer,
              CircleEnd = !sourceLeader.Dimldrblk.IsNull,
            };
            for (int vertex = 0; vertex < sourceLeader.NumVertices; vertex++)
            {
              captured.Vertices.Add(sourceLeader.VertexAt(vertex));
            }
            leaders.Add(captured);
            continue;
          }

          BlockReference block = entity as BlockReference;
          if (block == null)
          {
            continue;
          }

          string blockName = ResolveTemplateBlockName(transaction, block);
          if (string.IsNullOrEmpty(blockName) || blockName.StartsWith("*", StringComparison.Ordinal))
          {
            continue;
          }

          TemplateMember member = new TemplateMember
          {
            BlockName = blockName,
            Layer = block.Layer,
            ScaleRatio = block.ScaleFactors.X / blockScale,
          };
          ReadTemplateAttributes(transaction, block, member);

          if (string.Equals(blockName, KnBlockName, StringComparison.OrdinalIgnoreCase))
          {
            member.Kind = TemplateMemberKind.Note;
            member.RelativeRotation = RoomTemplateGeometry.NormalizeAngle(block.Rotation - ucsRotation);
            notes.Add(member);
            notePositions.Add(block.Position);
            continue;
          }

          Vector3d local = RoomTemplateGeometry.ToLocal(block.Position, origin, ucsRotation);
          member.X = local.X;
          member.Y = local.Y;
          member.RelativeRotation = RoomTemplateGeometry.NormalizeAngle(block.Rotation - ucsRotation);

          if (IsSupportedReceptacleName(blockName))
          {
            member.Kind = TemplateMemberKind.Receptacle;
            member.VisibilityState = ResolveReceptacleVisibilityState(block);
            groupReceptacles.Add(
              new CapturedReceptacle { ObjectId = id, Position = block.Position, Member = member }
            );
          }
          else
          {
            member.Kind = TemplateMemberKind.Symbol;
            ReadTemplateDynamicProperties(block, member);
          }
          symbols.Add(member);
          symbolPositions.Add(block.Position);
        }
        transaction.Commit();
      }

      group.Members.AddRange(symbols);
      int firstNote = group.Members.Count;
      group.Members.AddRange(notes);
      int firstLeader = group.Members.Count;

      bool[] noteTaken = new bool[notes.Count];
      for (int leaderIndex = 0; leaderIndex < leaders.Count; leaderIndex++)
      {
        CapturedLeader leader = leaders[leaderIndex];
        if (leader.Vertices.Count < 2)
        {
          continue;
        }

        Point3d tip = leader.Vertices[0];
        Point3d end = leader.Vertices[leader.Vertices.Count - 1];
        int parent = FindNearestIndex(symbolPositions, tip);
        Point3d anchor = parent >= 0 ? symbolPositions[parent] : origin;

        TemplateMember member = new TemplateMember
        {
          Kind = TemplateMemberKind.Leader,
          Layer = leader.Layer,
          ParentIndex = parent,
          CircleEnd = leader.CircleEnd,
        };
        Vector3d tipOffset = ToPlottedOffset(tip, anchor, ucsRotation, blockScale);
        member.X = tipOffset.X;
        member.Y = tipOffset.Y;
        foreach (Point3d vertex in leader.Vertices)
        {
          Vector3d vertexOffset = ToPlottedOffset(vertex, anchor, ucsRotation, blockScale);
          member.LeaderPath.Add(new TemplatePoint { X = vertexOffset.X, Y = vertexOffset.Y });
        }

        int noteIndex = FindNearestIndex(
          notePositions,
          end,
          TemplateNoteMatchPlottedInches * blockScale,
          noteTaken
        );
        if (noteIndex >= 0)
        {
          noteTaken[noteIndex] = true;
          member.NoteIndex = firstNote + noteIndex;
          notes[noteIndex].ParentIndex = parent;
        }
        group.Members.Add(member);
      }

      for (int noteIndex = 0; noteIndex < notes.Count; noteIndex++)
      {
        TemplateMember note = notes[noteIndex];
        if (!noteTaken[noteIndex])
        {
          note.ParentIndex = FindNearestIndex(symbolPositions, notePositions[noteIndex]);
        }

        Point3d anchor = note.ParentIndex >= 0 ? symbolPositions[note.ParentIndex] : origin;
        Vector3d offset = ToPlottedOffset(notePositions[noteIndex], anchor, ucsRotation, blockScale);
        note.X = offset.X;
        note.Y = offset.Y;
      }

      // Receptacles keep the order they have in the group so slot lookups stay stable.
      receptacles.AddRange(groupReceptacles);
      return group;
    }

    private static Vector3d ToPlottedOffset(
      Point3d point,
      Point3d anchor,
      double ucsRotation,
      double blockScale
    )
    {
      return RoomTemplateGeometry.ToLocal(point, anchor, ucsRotation) / blockScale;
    }

    // Index of the closest point, or -1 when none is within maxDistance.
    private static int FindNearestIndex(
      List<Point3d> points,
      Point3d target,
      double maxDistance = double.MaxValue,
      bool[] taken = null
    )
    {
      int best = -1;
      double bestDistance = maxDistance;
      for (int index = 0; index < points.Count; index++)
      {
        if (taken != null && taken[index])
        {
          continue;
        }
        double distance = points[index].DistanceTo(target);
        if (distance <= bestDistance)
        {
          bestDistance = distance;
          best = index;
        }
      }
      return best;
    }

    // Dynamic blocks are inserted as anonymous copies, so the real name comes from the
    // dynamic definition.
    private static string ResolveTemplateBlockName(Transaction transaction, BlockReference block)
    {
      ObjectId definitionId = block.IsDynamicBlock ? block.DynamicBlockTableRecord : block.BlockTableRecord;
      BlockTableRecord definition = transaction.GetObject(definitionId, OpenMode.ForRead) as BlockTableRecord;
      return definition?.Name ?? string.Empty;
    }

    private static void ReadTemplateAttributes(
      Transaction transaction,
      BlockReference block,
      TemplateMember member
    )
    {
      foreach (ObjectId attributeId in block.AttributeCollection)
      {
        AttributeReference attribute = transaction.GetObject(attributeId, OpenMode.ForRead) as AttributeReference;
        if (attribute != null)
        {
          member.Attributes[attribute.Tag.ToUpperInvariant()] = attribute.TextString;
        }
      }
    }

    private static void ReadTemplateDynamicProperties(BlockReference block, TemplateMember member)
    {
      if (!block.IsDynamicBlock)
      {
        return;
      }

      foreach (DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection)
      {
        if (property.ReadOnly)
        {
          continue;
        }

        object value = property.Value;
        if (value is string || value is double || value is int || value is short)
        {
          member.DynamicProperties[property.PropertyName] = Convert.ToString(
            value,
            CultureInfo.InvariantCulture
          );
        }
      }
    }

    // Decides which receptacles share a circuit: by the circuit label text next to them
    // when the user selects it, otherwise by the receptacles the user picks per circuit.
    private static bool TryAssignTemplateCircuits(
      Database db,
      Editor ed,
      RoomTemplate template,
      List<CapturedReceptacle> receptacles
    )
    {
      PromptSelectionResult labelSelection = ed.GetSelection(
        new PromptSelectionOptions
        {
          MessageForAdding =
            "\nSelect the circuit label text in the room, or press Enter to pick the receptacles on each circuit instead: ",
        },
        new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT") })
      );

      if (labelSelection.Status == PromptStatus.OK && labelSelection.Value != null)
      {
        AssignCircuitsFromLabels(db, template, receptacles, labelSelection.Value.GetObjectIds());
        int unlabeled = 0;
        foreach (CapturedReceptacle receptacle in receptacles)
        {
          if (receptacle.Member.CircuitSlot < 0)
          {
            unlabeled++;
          }
        }
        if (unlabeled > 0)
        {
          ed.WriteMessage(
            $"\n{unlabeled} receptacle(s) have no circuit label nearby and are not on a circuit yet."
          );
        }
        if (unlabeled == receptacles.Count)
        {
          ed.WriteMessage("\nNo circuit labels matched any receptacle.");
        }
        else if (unlabeled == 0)
        {
          WarnAboutOverloadedTemplateCircuits(db, ed, template, receptacles);
          return true;
        }
      }

      if (!PickTemplateCircuits(db, ed, template, receptacles))
      {
        return false;
      }
      WarnAboutOverloadedTemplateCircuits(db, ed, template, receptacles);
      return true;
    }

    private static void AssignCircuitsFromLabels(
      Database db,
      RoomTemplate template,
      List<CapturedReceptacle> receptacles,
      ObjectId[] labelIds
    )
    {
      List<string> labels = new List<string>();
      List<Point3d> labelPositions = new List<Point3d>();
      using (Transaction transaction = db.TransactionManager.StartOpenCloseTransaction())
      {
        foreach (ObjectId id in labelIds)
        {
          Entity entity = transaction.GetObject(id, OpenMode.ForRead, false) as Entity;
          string text = (entity as MText)?.Text ?? (entity as DBText)?.TextString;
          text = Regex.Replace(text ?? string.Empty, @"\s+", " ").Trim().ToUpperInvariant();
          if (text.Length == 0)
          {
            continue;
          }

          Point3d position;
          try
          {
            Extents3d extents = entity.GeometricExtents;
            position = new Point3d(
              (extents.MinPoint.X + extents.MaxPoint.X) / 2.0,
              (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0,
              0.0
            );
          }
          catch
          {
            position = (entity as MText)?.Location ?? (entity as DBText).Position;
          }
          labels.Add(text);
          labelPositions.Add(position);
        }
      }

      double cutoff = TemplateLabelCutoffPlottedInches * template.CaptureBlockScale;
      List<string> distinct = new List<string>();
      string[] assigned = new string[receptacles.Count];
      for (int index = 0; index < receptacles.Count; index++)
      {
        Point3d flat = new Point3d(receptacles[index].Position.X, receptacles[index].Position.Y, 0.0);
        int nearest = FindNearestIndex(labelPositions, flat, cutoff);
        if (nearest < 0)
        {
          continue;
        }
        assigned[index] = labels[nearest];
        if (!distinct.Contains(labels[nearest]))
        {
          distinct.Add(labels[nearest]);
        }
      }

      // Circuit 3 sorts before circuit 12, whatever the panel prefix.
      distinct.Sort(CompareCircuitLabels);
      for (int index = 0; index < receptacles.Count; index++)
      {
        if (assigned[index] != null)
        {
          receptacles[index].Member.CircuitSlot = template.CircuitLabels.Count + distinct.IndexOf(assigned[index]);
        }
      }
      template.CircuitLabels.AddRange(distinct);
    }

    private static int CompareCircuitLabels(string first, string second)
    {
      Match firstMatch = Regex.Match(first, @"(\d+)\s*$");
      Match secondMatch = Regex.Match(second, @"(\d+)\s*$");
      if (firstMatch.Success && secondMatch.Success)
      {
        int byNumber = int.Parse(firstMatch.Groups[1].Value, CultureInfo.InvariantCulture)
          .CompareTo(int.Parse(secondMatch.Groups[1].Value, CultureInfo.InvariantCulture));
        if (byNumber != 0)
        {
          return byNumber;
        }
      }
      return string.Compare(first, second, StringComparison.OrdinalIgnoreCase);
    }

    // Lets the user pick the receptacles on circuit 1, then circuit 2, and so on, for any
    // receptacle that is not on a circuit yet. Enter on an empty pick ends the loop.
    private static bool PickTemplateCircuits(
      Database db,
      Editor ed,
      RoomTemplate template,
      List<CapturedReceptacle> receptacles
    )
    {
      Dictionary<ObjectId, CapturedReceptacle> unassigned = new Dictionary<ObjectId, CapturedReceptacle>();
      foreach (CapturedReceptacle receptacle in receptacles)
      {
        if (receptacle.Member.CircuitSlot < 0)
        {
          unassigned[receptacle.ObjectId] = receptacle;
        }
      }

      SelectionFilter filter = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "INSERT") });
      while (unassigned.Count > 0)
      {
        int slot = template.CircuitLabels.Count;
        PromptSelectionResult pick = ed.GetSelection(
          new PromptSelectionOptions
          {
            MessageForAdding =
              $"\nSelect the receptacles on circuit {slot + 1} ({unassigned.Count} left, Enter when done): ",
          },
          filter
        );
        if (pick.Status != PromptStatus.OK || pick.Value == null)
        {
          break;
        }

        int count = 0;
        foreach (ObjectId id in pick.Value.GetObjectIds())
        {
          if (unassigned.TryGetValue(id, out CapturedReceptacle receptacle))
          {
            receptacle.Member.CircuitSlot = slot;
            unassigned.Remove(id);
            count++;
          }
        }
        if (count == 0)
        {
          ed.WriteMessage("\nNone of those are receptacles in the template that still need a circuit.");
          continue;
        }
        template.CircuitLabels.Add($"Circuit {slot + 1}");
      }

      if (unassigned.Count > 0)
      {
        ed.WriteMessage($"\n{unassigned.Count} receptacle(s) will not be circuited by this template.");
      }
      return true;
    }

    private static void WarnAboutOverloadedTemplateCircuits(
      Database db,
      Editor ed,
      RoomTemplate template,
      List<CapturedReceptacle> receptacles
    )
    {
      double maximumKva = ResolveReceptacleCircuitMaximumKva(db);
      for (int slot = 0; slot < template.CircuitLabels.Count; slot++)
      {
        List<ObjectId> ids = new List<ObjectId>();
        foreach (CapturedReceptacle receptacle in receptacles)
        {
          if (receptacle.Member.CircuitSlot == slot)
          {
            ids.Add(receptacle.ObjectId);
          }
        }
        if (ids.Count == 0)
        {
          continue;
        }

        double kva = SumReceptacleLoadUnits(db, ids) * ReceptacleLoadUnitKva;
        if (kva > maximumKva + 1e-9)
        {
          ed.WriteMessage(
            $"\nWarning: {template.CircuitLabels[slot]} carries {kva:0.00} kVA, above the "
              + $"{maximumKva:0.00} kVA circuit maximum set in HRS. The template keeps the split as drawn."
          );
        }
      }
    }

    private static int SumReceptacleLoadUnits(Database db, List<ObjectId> ids)
    {
      int units = 0;
      foreach (ReceptacleLoadItem item in ReadReceptacleLoadItems(db, ids.ToArray(), out _, out _, out _))
      {
        units += item.LoadUnits;
      }
      return units;
    }

    [CommandMethod("RTLIST", CommandFlags.Modal)]
    [CommandMethod("RTD", CommandFlags.Modal)]
    public static void ListRoomTemplates()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      List<string> names = RoomTemplateStore.ListNames(db);
      if (names.Count == 0)
      {
        ed.WriteMessage("\nThis drawing has no room templates. Run RTC to capture one.");
        return;
      }

      ed.WriteMessage("\nRoom templates in this drawing:");
      foreach (string name in names)
      {
        if (RoomTemplateStore.TryLoad(db, name, out RoomTemplate template))
        {
          ed.WriteMessage($"\n  {template.Name}: {DescribeRoomTemplate(template)}");
        }
      }

      PromptStringOptions options = new PromptStringOptions(
        "\nEnter a template name to delete, or press Enter to leave them all: "
      )
      {
        AllowSpaces = true,
      };
      PromptResult result = ed.GetString(options);
      string toDelete = NormalizeTemplateName(result.Status == PromptStatus.OK ? result.StringResult : string.Empty);
      if (toDelete.Length == 0)
      {
        return;
      }

      try
      {
        ed.WriteMessage(
          RoomTemplateStore.Delete(db, toDelete)
            ? $"\nDeleted room template {toDelete}."
            : $"\nNo room template named {toDelete}."
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to delete the template: {ex.Message}");
      }
    }

    private static string DescribeRoomTemplate(RoomTemplate template)
    {
      int symbols = 0;
      foreach (TemplateGroup group in template.Groups)
      {
        foreach (TemplateMember member in group.Members)
        {
          if (member.IsSymbol)
          {
            symbols++;
          }
        }
      }
      return $"{template.Groups.Count} group(s), {symbols} symbol(s), {template.CircuitLabels.Count} circuit(s)";
    }
  }
}
