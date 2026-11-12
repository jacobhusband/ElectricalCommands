using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private sealed class PlacedTemplateGroup
    {
      internal List<ObjectId> Entities { get; } = new List<ObjectId>();

      // Circuit slot and object id of each receptacle in the group.
      internal List<KeyValuePair<int, ObjectId>> Receptacles { get; } =
        new List<KeyValuePair<int, ObjectId>>();
    }

    [CommandMethod("RTPLACE", CommandFlags.Modal)]
    [CommandMethod("RTP", CommandFlags.Modal)]
    public static void PlaceRoomTemplate()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      List<string> names = RoomTemplateStore.ListNames(db);
      if (names.Count == 0)
      {
        ed.WriteMessage("\nThis drawing has no room templates. Draw one room, then run RTC to capture it.");
        return;
      }

      if (!TryResolveTemplateScale(db, ed, "RTP", out var scale)) return;
      double blockScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
      double ucsRotation = SymbolGeometry.ResolveUcsRotation(ed);

      if (!TryChooseRoomTemplate(db, ed, names, out RoomTemplate template)) return;

      int slotCount = template.CircuitLabels.Count;
      foreach (TemplateGroup group in template.Groups)
      {
        foreach (TemplateMember member in group.Members)
        {
          slotCount = Math.Max(slotCount, member.CircuitSlot + 1);
        }
      }

      if (!TryResolveTemplateBlocks(db, ed, template, out Dictionary<string, ObjectId> definitions))
      {
        return;
      }

      ed.WriteMessage(
        $"\nPlacing room template {template.Name} ({template.Groups.Count} group(s)). "
          + "Space/Enter rotates 90 degrees, Wall aligns to a wall, Skip leaves a group out, "
          + "Back redoes the previous group, Done finishes, Esc stops."
      );

      PlacedTemplateGroup[] placed = new PlacedTemplateGroup[template.Groups.Count];
      double userRotation = 0.0;
      int index = 0;
      bool stopped = false;
      bool finished = false;

      while (index < template.Groups.Count && !stopped && !finished)
      {
        TemplateGroup group = template.Groups[index];
        bool advance = false;
        int draggingIndex = index;
        List<ResolvedTemplateMember> placement = null;
        // Dynamic properties and attributes require database-resident previews.
        // Keep them in an anonymous definition, then erase them and commit the
        // cleanup. Aborting this transaction also rewinds pan/zoom performed
        // during the jig, which makes the camera jump between group placements.
        using (Transaction previewTransaction = db.TransactionManager.StartTransaction())
        {
          var previewTable = (BlockTable)previewTransaction.GetObject(db.BlockTableId, OpenMode.ForWrite);
          var previewSpace = new BlockTableRecord { Name = "*U" };
          previewTable.Add(previewSpace);
          previewTransaction.AddNewlyCreatedDBObject(previewSpace, true);
          Func<TemplateMember, BlockReference> createPreview = member =>
          {
            if (!definitions.TryGetValue(TemplateDefinitionKey(member), out ObjectId definitionId)) return null;
            return CreateTemplateBlock(db, previewTransaction, previewSpace, definitionId,
              new ResolvedTemplateMember { Member = member, Position = Point3d.Origin,
                Rotation = 0, ScaleFactors = new Scale3d(1) });
          };
          using (
            RoomTemplateGroupJig jig = new RoomTemplateGroupJig(
              db,
              group,
              createPreview,
              blockScale,
              ucsRotation,
              userRotation
            )
          )
          {
            while (!advance && !stopped)
            {
              jig.Message =
                $"\nPlace group {index + 1} of {template.Groups.Count} "
                + "[Rotate/Flip/Wall/Angle/Skip/Back/Done] "
                + $"<{Math.Round(userRotation * 180.0 / Math.PI)}°>: ";
              PromptResult result = ed.Drag(jig);
              if (
                result.Status != PromptStatus.OK
                && result.Status != PromptStatus.None
                && result.Status != PromptStatus.Keyword
              )
              {
                stopped = true;
                break;
              }

              switch (jig.LastInput)
              {
                case TemplateJigInput.Point:
                  try
                  {
                    placement = RoomTemplateGeometry.Resolve(
                      group,
                      jig.Location,
                      ucsRotation,
                      userRotation,
                      blockScale
                    );
                    advance = true;
                  }
                  catch (System.Exception ex)
                  {
                    ed.WriteMessage($"\nUnable to place group {index + 1}: {ex.Message}");
                    stopped = true;
                  }
                  break;

                case TemplateJigInput.Enter:
                  userRotation = RoomTemplateGeometry.NormalizeAngle(userRotation + Math.PI / 2.0);
                  jig.UserRotation = userRotation;
                  break;

                case TemplateJigInput.Keyword:
                  HandleTemplateKeyword(
                    ed,
                    jig.LastKeyword,
                    group,
                    ref index,
                    ref userRotation,
                    ref advance,
                    ref finished,
                    ucsRotation
                  );
                  jig.UserRotation = userRotation;
                  if (finished)
                  {
                    advance = true;
                  }
                  break;

                default:
                  stopped = true;
                  break;
              }
            }
          }
          CleanupTemplatePreview(previewTransaction, previewSpace);
          previewTransaction.Commit();
        }
        if (index < draggingIndex)
        {
          EraseTemplateGroup(db, placed[index]);
          placed[index] = null;
        }
        if (placement != null)
        {
          try
          {
            placed[index] = InsertTemplateGroup(db, placement, definitions, blockScale);
            index++;
          }
          catch (System.Exception ex)
          {
            ed.WriteMessage($"\nUnable to place group {index + 1}: {ex.Message}");
            stopped = true;
          }
        }
      }

      int placedGroups = 0;
      foreach (PlacedTemplateGroup group in placed)
      {
        if (group != null)
        {
          placedGroups++;
        }
      }

      if (stopped)
      {
        ed.WriteMessage(
          placedGroups == 0
            ? "\nRTP canceled."
            : $"\nRTP stopped after placing {placedGroups} group(s); nothing was circuited. "
              + "Select the receptacles and run RC to circuit them."
        );
        return;
      }

      List<ObjectId>[] slotReceptacles = new List<ObjectId>[slotCount];
      for (int slot = 0; slot < slotCount; slot++)
      {
        slotReceptacles[slot] = new List<ObjectId>();
      }
      foreach (PlacedTemplateGroup group in placed)
      {
        if (group == null)
        {
          continue;
        }
        foreach (KeyValuePair<int, ObjectId> receptacle in group.Receptacles)
        {
          if (receptacle.Key >= 0)
          {
            slotReceptacles[receptacle.Key].Add(receptacle.Value);
          }
        }
      }

      ed.WriteMessage($"\nPlaced {placedGroups} of {template.Groups.Count} group(s) from {template.Name}.");
      if (slotCount == 0 || CountSlotReceptacles(slotReceptacles) == 0)
      {
        return;
      }
      PromptKeywordOptions circuitOptions = new PromptKeywordOptions(
        "\nName and circuit this room now? [Yes/No] <Yes>: ", "Yes No")
      {
        AllowNone = true,
      };
      circuitOptions.Keywords.Default = "Yes";
      PromptResult circuitResult = ed.GetKeywords(circuitOptions);
      if ((circuitResult.Status != PromptStatus.OK && circuitResult.Status != PromptStatus.None)
        || string.Equals(circuitResult.StringResult, "No", StringComparison.OrdinalIgnoreCase))
      {
        ed.WriteMessage("\nRoom naming and circuiting skipped. Select the receptacles and run RC when ready.");
        return;
      }
      if (!TryPrepareTemplateCircuiting(db, ed, out string panelName,
        out ElectricalDrawingSettingsStore.PanelScheduleSetting panelSchedule))
      {
        ed.WriteMessage("\nThe room remains placed. Select the receptacles and run RC when ready.");
        return;
      }
      CircuitTemplateRoom(db, ed, template, slotReceptacles, scale, panelName, panelSchedule);
    }

    internal static void CleanupTemplatePreview(Transaction transaction, BlockTableRecord previewSpace)
    {
      foreach (ObjectId id in previewSpace)
      {
        DBObject entity = transaction.GetObject(id, OpenMode.ForWrite);
        if (!entity.IsErased) entity.Erase();
      }
      previewSpace.Erase();
    }

    private static void HandleTemplateKeyword(
      Editor ed,
      string keyword,
      TemplateGroup group,
      ref int index,
      ref double userRotation,
      ref bool advance,
      ref bool finished,
      double ucsRotation
    )
    {
      switch (keyword)
      {
        case "Rotate":
          userRotation = RoomTemplateGeometry.NormalizeAngle(userRotation + Math.PI / 2.0);
          break;

        case "Flip":
          userRotation = RoomTemplateGeometry.NormalizeAngle(userRotation + Math.PI);
          break;

        case "Wall":
          if (TryPickWallAngle(ed, out double wallAngle))
          {
            // The group faces the left-hand side of the wall as picked; Flip turns it around.
            userRotation = RoomTemplateGeometry.RotationForWall(
              group,
              wallAngle,
              RoomTemplateGeometry.NormalizeAngle(wallAngle + Math.PI / 2.0),
              ucsRotation
            );
          }
          break;

        case "Angle":
          PromptDoubleOptions angleOptions = new PromptDoubleOptions(
            "\nEnter the rotation in degrees (0 = as captured) "
          )
          {
            DefaultValue = Math.Round(userRotation * 180.0 / Math.PI),
            UseDefaultValue = true,
          };
          PromptDoubleResult angleResult = ed.GetDouble(angleOptions);
          if (angleResult.Status == PromptStatus.OK)
          {
            userRotation = RoomTemplateGeometry.NormalizeAngle(angleResult.Value * Math.PI / 180.0);
          }
          break;

        case "Skip":
          index++;
          advance = true;
          break;

        case "Back":
          if (index == 0)
          {
            ed.WriteMessage("\nThere is no earlier group to redo.");
            break;
          }
          index--;
          advance = true;
          break;

        case "Done":
          finished = true;
          break;
      }
    }

    private static bool TryPickWallAngle(Editor ed, out double wallAngle)
    {
      wallAngle = 0.0;
      PromptPointResult first = ed.GetPoint("\nPick the first point along the wall: ");
      if (first.Status != PromptStatus.OK)
      {
        return false;
      }

      PromptPointOptions secondOptions = new PromptPointOptions("\nPick the second point along the wall: ")
      {
        BasePoint = first.Value,
        UseBasePoint = true,
      };
      PromptPointResult second = ed.GetPoint(secondOptions);
      if (second.Status != PromptStatus.OK)
      {
        return false;
      }

      Vector3d direction = (second.Value - first.Value).TransformBy(ed.CurrentUserCoordinateSystem);
      if (Math.Sqrt(direction.X * direction.X + direction.Y * direction.Y) < 1e-6)
      {
        ed.WriteMessage("\nThe two wall points are the same point.");
        return false;
      }

      wallAngle = Math.Atan2(direction.Y, direction.X);
      return true;
    }

    private static bool TryChooseRoomTemplate(
      Database db,
      Editor ed,
      List<string> names,
      out RoomTemplate template
    )
    {
      template = null;
      string defaultName = names[0];
      foreach (string name in names)
      {
        if (string.Equals(name, _lastRoomTemplateName, StringComparison.OrdinalIgnoreCase))
        {
          defaultName = name;
        }
      }

      string chosen = defaultName;
      if (names.Count > 1)
      {
        while (true)
        {
          PromptStringOptions options = new PromptStringOptions(
            $"\nRoom template [{string.Join("/", names)}] <{defaultName}>: "
          )
          {
            AllowSpaces = true,
          };
          PromptResult result = ed.GetString(options);
          if (result.Status != PromptStatus.OK && result.Status != PromptStatus.None)
          {
            ed.WriteMessage("\nRTP canceled.");
            return false;
          }

          string typed = NormalizeTemplateName(result.StringResult);
          if (typed.Length == 0)
          {
            break;
          }
          if (names.Exists(name => string.Equals(name, typed, StringComparison.OrdinalIgnoreCase)))
          {
            chosen = typed;
            break;
          }
          ed.WriteMessage($"\nNo room template named {typed}.");
        }
      }

      if (!RoomTemplateStore.TryLoad(db, chosen, out template))
      {
        ed.WriteMessage($"\nUnable to read room template {chosen}.");
        return false;
      }
      _lastRoomTemplateName = template.Name;
      return true;
    }

    // The same checks RC makes, deferred until the user chooses to circuit the
    // placed room. Placement alone needs no panel name or linked schedule.
    private static bool TryPrepareTemplateCircuiting(
      Database db,
      Editor ed,
      out string panelName,
      out ElectricalDrawingSettingsStore.PanelScheduleSetting panelSchedule
    )
    {
      panelSchedule = null;
      if (!IsInModelOrViewportSpace(db))
      {
        panelName = string.Empty;
        ed.WriteMessage(
          "\nRTP circuits receptacles, so run it from model space or from inside an active "
            + "paper-space viewport. Double-click inside a viewport and try again."
        );
        return false;
      }

      if (!ElectricalDrawingSettingsStore.TryReadPanelName(db, out panelName))
      {
        ed.WriteMessage("\nRTP requires a panel name to circuit the receptacles. Run SETPANELNAME (SPN) first.");
        return false;
      }

      bool hasSchedule =
        ElectricalDrawingSettingsStore.TryReadPanelSchedule(db, out panelSchedule)
        && System.IO.File.Exists(panelSchedule.WorkbookPath);
      if (!hasSchedule)
      {
        ed.WriteMessage("\nRTP requires a linked panel schedule. Select the panel schedule now.");
        SetPanelScheduleCommand();
        hasSchedule =
          ElectricalDrawingSettingsStore.TryReadPanelSchedule(db, out panelSchedule)
          && System.IO.File.Exists(panelSchedule.WorkbookPath);
        if (!hasSchedule)
        {
          ed.WriteMessage("\nRTP circuiting skipped because no usable panel schedule was linked.");
          return false;
        }
      }

      if (!TryVerifyPanelScheduleWorkbookClosed(panelSchedule.WorkbookPath, out string availabilityError))
      {
        ed.WriteMessage($"\nRTP circuiting skipped: {availabilityError.Replace("running RC", "running RTP")}");
        return false;
      }
      return true;
    }

    private static string TemplateDefinitionKey(TemplateMember member)
    {
      return member.Kind == TemplateMemberKind.Note ? KnBlockName : member.BlockName.ToUpperInvariant();
    }

    // Finds (or, for the plugin's own symbols, rebuilds) every block the template uses.
    private static bool TryResolveTemplateBlocks(
      Database db,
      Editor ed,
      RoomTemplate template,
      out Dictionary<string, ObjectId> definitions
    )
    {
      definitions = new Dictionary<string, ObjectId>();
      bool usesKeyedNotes = false;
      foreach (TemplateGroup group in template.Groups)
      {
        foreach (TemplateMember member in group.Members)
        {
          usesKeyedNotes |= member.Kind == TemplateMemberKind.Note || member.Kind == TemplateMemberKind.Leader;
        }
      }

      try
      {
        if (usesKeyedNotes)
        {
          EnsureKnLayer(db);
          ObjectId textStyleId = EnsureKnTextStyle(db);
          if (!EnsureKnBlockDefinition(db, textStyleId))
          {
            ed.WriteMessage($"\nFailed to prepare the {KnBlockName} block definition.");
            return false;
          }
        }

        List<string> missing = new List<string>();
        using (Transaction transaction = db.TransactionManager.StartTransaction())
        {
          BlockTable blockTable = (BlockTable)transaction.GetObject(db.BlockTableId, OpenMode.ForRead);
          foreach (TemplateGroup group in template.Groups)
          {
            foreach (TemplateMember member in group.Members)
            {
              if (member.Kind == TemplateMemberKind.Leader)
              {
                continue;
              }

              string key = TemplateDefinitionKey(member);
              if (definitions.ContainsKey(key))
              {
                continue;
              }

              ObjectId definitionId = ObjectId.Null;
              string name = member.Kind == TemplateMemberKind.Note ? KnBlockName : member.BlockName;
              if (blockTable.Has(name))
              {
                definitionId = blockTable[name];
              }
              else if (member.Kind == TemplateMemberKind.Receptacle)
              {
                TryResolveReceptBlockDefinition(blockTable, out definitionId, out _);
              }
              else if (!TryCreateKnownSymbolBlock(transaction, db, name, out definitionId))
              {
                definitionId = ObjectId.Null;
              }

              if (definitionId.IsNull)
              {
                if (!missing.Contains(name))
                {
                  missing.Add(name);
                }
                continue;
              }
              definitions[key] = definitionId;
            }
          }
          transaction.Commit();
        }

        if (missing.Count > 0)
        {
          ed.WriteMessage(
            $"\nThe template needs block(s) this drawing does not have: {string.Join(", ", missing)}. "
              + "Copy one of each into this drawing and run RTP again."
          );
          return false;
        }
        return true;
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to prepare the template blocks: {ex.Message}");
        return false;
      }
    }

    // The plugin's own symbol commands create their blocks on first use, so a drawing
    // that has not placed one yet can still get it from here.
    private static bool TryCreateKnownSymbolBlock(
      Transaction transaction,
      Database db,
      string blockName,
      out ObjectId definitionId
    )
    {
      definitionId = ObjectId.Null;
      Func<List<Entity>> createEntities = null;
      switch ((blockName ?? string.Empty).ToUpperInvariant())
      {
        case "DATA":
          createEntities = () => DataGeometry.CreateEntities(db, SymbolSide.South);
          break;
        case "ACIES_DATA":
          createEntities = () => DataGeometry.CreateEntities(db, SymbolSide.North);
          break;
        case "DISCONNECT":
          createEntities = () => DisconnectGeometry.CreateEntities(db, SymbolSide.North);
          break;
        case "SWITCH":
          createEntities = () => SwitchSymbolGeometry.CreateEntities(db, SwitchSymbolStyle.Standard);
          break;
        case "SWITCH-DIMMER":
          createEntities = () => SwitchSymbolGeometry.CreateEntities(db, SwitchSymbolStyle.Dimmer);
          break;
        case "SWITCH-OCCUPANCY":
          createEntities = () => SwitchSymbolGeometry.CreateEntities(db, SwitchSymbolStyle.Occupancy);
          break;
        case "JBOX":
          createEntities = () =>
            JunctionBoxGeometry.CreateEntities(db, JunctionBoxStyle.Plain, SymbolSide.North);
          break;
        default:
          Match accessory = Regex.Match(
            blockName ?? string.Empty,
            @"^JBOX-(SWITCH|WALLMOUNT)-([NESW])$",
            RegexOptions.IgnoreCase
          );
          if (accessory.Success)
          {
            JunctionBoxStyle style = string.Equals(
              accessory.Groups[1].Value,
              "SWITCH",
              StringComparison.OrdinalIgnoreCase
            )
              ? JunctionBoxStyle.Switch
              : JunctionBoxStyle.WallMount;
            SymbolSide side = SymbolSide.North;
            switch (char.ToUpperInvariant(accessory.Groups[2].Value[0]))
            {
              case 'E':
                side = SymbolSide.East;
                break;
              case 'S':
                side = SymbolSide.South;
                break;
              case 'W':
                side = SymbolSide.West;
                break;
            }
            createEntities = () => JunctionBoxGeometry.CreateEntities(db, style, side);
          }
          break;
      }

      if (createEntities == null)
      {
        return false;
      }
      definitionId = SymbolGeometry.EnsureBlock(transaction, db, blockName.ToUpperInvariant(), createEntities);
      return true;
    }

    private static PlacedTemplateGroup InsertTemplateGroup(
      Database db,
      List<ResolvedTemplateMember> resolved,
      Dictionary<string, ObjectId> definitions,
      double blockScale
    )
    {
      PlacedTemplateGroup placed = new PlacedTemplateGroup();
      using (Transaction transaction = db.TransactionManager.StartTransaction())
      {
        BlockTableRecord space = (BlockTableRecord)transaction.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

        foreach (ResolvedTemplateMember item in resolved)
        {
          TemplateMember member = item.Member;
          if (member.Kind == TemplateMemberKind.Leader)
          {
            if (item.LeaderPath == null || item.LeaderPath.Length < 2)
            {
              continue;
            }

            Leader leader = CreateKnLeader(
              db,
              new KeyNoteLeaderLayout(Point3d.Origin, item.LeaderPath),
              member.CircleEnd ? KeyNoteLeaderEnd.Circle : KeyNoteLeaderEnd.Arrow,
              blockScale
            );
            space.AppendEntity(leader);
            transaction.AddNewlyCreatedDBObject(leader, true);
            placed.Entities.Add(leader.ObjectId);
            continue;
          }

          if (!definitions.TryGetValue(TemplateDefinitionKey(member), out ObjectId definitionId))
          {
            continue;
          }

          BlockReference block = CreateTemplateBlock(db, transaction, space, definitionId, item);

          placed.Entities.Add(block.ObjectId);
          if (member.Kind == TemplateMemberKind.Receptacle)
          {
            placed.Receptacles.Add(new KeyValuePair<int, ObjectId>(member.CircuitSlot, block.ObjectId));
          }
        }
        transaction.Commit();
      }
      return placed;
    }

    // One setup path for the jig and final insertion: same definition, dynamic
    // state, signed scales, layer and attribute geometry/value.
    private static BlockReference CreateTemplateBlock(Database db, Transaction transaction,
      BlockTableRecord space, ObjectId definitionId, ResolvedTemplateMember item)
    {
      TemplateMember member = item.Member;
      var block = new BlockReference(item.Position, definitionId);
      block.SetDatabaseDefaults(db);
      block.ScaleFactors = item.ScaleFactors;
      block.Rotation = item.Rotation;
      var layers = (LayerTable)transaction.GetObject(db.LayerTableId, OpenMode.ForRead);
      string layer = member.Kind == TemplateMemberKind.Note && !layers.Has(member.Layer) ? KnLayerName : member.Layer;
      if (!string.IsNullOrEmpty(layer) && layers.Has(layer)) block.Layer = layer;
      space.AppendEntity(block);
      transaction.AddNewlyCreatedDBObject(block, true);
      ApplyTemplateDynamicProperties(block, member);
      if (member.Kind == TemplateMemberKind.Receptacle && !string.IsNullOrWhiteSpace(member.VisibilityState)
        && !TrySetReceptacleVisibilityState(block, member.VisibilityState))
        throw new InvalidOperationException($"The {member.BlockName} block does not expose receptacle type {member.VisibilityState}.");

      AddDefaultAttributes(transaction, definitionId, block);
      UpdateTemplateBlock(db, transaction, block, item);
      return block;
    }

    internal static void UpdateTemplateBlock(Database db, Transaction transaction,
      BlockReference block, ResolvedTemplateMember item)
    {
      block.Position = item.Position;
      block.Rotation = item.Rotation;
      block.ScaleFactors = item.ScaleFactors;
      TemplateMember member = item.Member;
      var definition = (BlockTableRecord)transaction.GetObject(block.IsDynamicBlock
        ? block.DynamicBlockTableRecord : block.BlockTableRecord, OpenMode.ForRead);
      var attributeDefinitions = new Dictionary<string, AttributeDefinition>();
      foreach (ObjectId id in definition)
        if (transaction.GetObject(id, OpenMode.ForRead) is AttributeDefinition ad && !ad.Constant)
          attributeDefinitions[ad.Tag.ToUpperInvariant()] = ad;
      foreach (ObjectId attributeId in block.AttributeCollection)
      {
        var attribute = (AttributeReference)transaction.GetObject(attributeId, OpenMode.ForWrite);
        string tag = attribute.Tag.ToUpperInvariant();
        if (attributeDefinitions.TryGetValue(tag, out AttributeDefinition ad))
          attribute.SetAttributeFromBlock(ad, block.BlockTransform);
        TemplateAttribute geometry = member.AttributeGeometry.Find(candidate => candidate.Tag == tag);
        if (geometry != null)
        {
          // Start from local geometry; applying the complete block transform also
          // restores MIRRTEXT corrections stored on the original attribute.
          attribute.TransformBy(block.BlockTransform.Inverse());
          attribute.HorizontalMode = (TextHorizontalMode)geometry.HorizontalMode;
          attribute.VerticalMode = (TextVerticalMode)geometry.VerticalMode;
          attribute.Position = new Point3d(geometry.Position);
          if (!attribute.IsDefaultAlignment) attribute.AlignmentPoint = new Point3d(geometry.AlignmentPoint);
          attribute.Rotation = geometry.Rotation;
          attribute.Height = geometry.Height;
          attribute.WidthFactor = geometry.WidthFactor;
          attribute.Oblique = geometry.Oblique;
          attribute.IsMirroredInX = geometry.MirroredX;
          attribute.IsMirroredInY = geometry.MirroredY;
          attribute.TransformBy(block.BlockTransform);
        }
        if (member.Attributes.TryGetValue(tag, out string text)) attribute.TextString = text;
        if (attribute.IsMTextAttribute) attribute.UpdateMTextAttribute();
        if (geometry == null)
          attribute.AdjustAlignment(db);
        else
        {
          // TransformBy can realign justified text. Preserve the independently
          // captured grip/alignment positions after transforming its text axes.
          if (!attribute.IsDefaultAlignment)
            attribute.AlignmentPoint = new Point3d(geometry.AlignmentPoint).TransformBy(block.BlockTransform);
          attribute.Position = new Point3d(geometry.Position).TransformBy(block.BlockTransform);
        }
      }
      block.RecordGraphicsModified(true);
    }

    private static void ApplyTemplateDynamicProperties(BlockReference block, TemplateMember member)
    {
      if (!block.IsDynamicBlock || member.DynamicProperties.Count == 0)
      {
        return;
      }

      foreach (DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection)
      {
        if (property.ReadOnly || !member.DynamicProperties.TryGetValue(property.PropertyName, out string text))
        {
          continue;
        }

        try
        {
          object current = property.Value;
          if (current is double)
          {
            property.Value = double.Parse(text, CultureInfo.InvariantCulture);
          }
          else if (current is int)
          {
            property.Value = int.Parse(text, CultureInfo.InvariantCulture);
          }
          else if (current is short)
          {
            property.Value = short.Parse(text, CultureInfo.InvariantCulture);
          }
          else
          {
            property.Value = text;
          }
        }
        catch
        {
          // A property this drawing's block does not accept keeps its default value.
        }
      }
    }

    private static void EraseTemplateGroup(Database db, PlacedTemplateGroup placed)
    {
      if (placed == null)
      {
        return;
      }

      using (Transaction transaction = db.TransactionManager.StartTransaction())
      {
        foreach (ObjectId id in placed.Entities)
        {
          DBObject entity = transaction.GetObject(id, OpenMode.ForWrite, true);
          if (!entity.IsErased)
          {
            entity.Erase();
          }
        }
        transaction.Commit();
      }
    }

    private static void CircuitTemplateRoom(
      Database db,
      Editor ed,
      RoomTemplate template,
      List<ObjectId>[] slotReceptacles,
      ElectricalDrawingSettingsStore.ScaleSetting scale,
      string panelName,
      ElectricalDrawingSettingsStore.PanelScheduleSetting panelSchedule
    )
    {
      List<int> slots = new List<int>();
      List<ObjectId> allReceptacles = new List<ObjectId>();
      for (int slot = 0; slot < slotReceptacles.Length; slot++)
      {
        if (slotReceptacles[slot].Count > 0)
        {
          slots.Add(slot);
          allReceptacles.AddRange(slotReceptacles[slot]);
        }
      }
      if (slots.Count == 0)
      {
        ed.WriteMessage("\nNo receptacles with a circuit were placed, so nothing was circuited.");
        return;
      }

      string detectedRoom = TryFindTaggedRoomName(db, CalculateReceptacleCenter(db, allReceptacles));
      if (!TryPromptTemplateRoomName(ed, detectedRoom, out string roomName))
      {
        ed.WriteMessage(
          "\nRTP placed the room but did not circuit it. Select the receptacles and run RC when ready."
        );
        return;
      }

      try
      {
        List<PanelScheduleCircuitRequest> requests = new List<PanelScheduleCircuitRequest>();
        foreach (int slot in slots)
        {
          requests.Add(
            new PanelScheduleCircuitRequest
            {
              ConnectedWatts = SumReceptacleLoadUnits(db, slotReceptacles[slot]) * ReceptacleLoadUnitKva * 1000.0,
              LoadDescription = "RECEPTACLES - " + roomName,
            }
          );
        }

        List<PanelScheduleAllocationResult> allocations =
          PanelScheduleWorkbookAllocator.AllocateReceptacleCircuits(
            panelSchedule.WorkbookPath,
            panelName,
            panelSchedule.CircuitCapacity,
            panelSchedule.SpareCount,
            requests
          );
        if (allocations.Count != requests.Count)
        {
          throw new InvalidOperationException("The panel schedule did not return every requested circuit.");
        }

        List<string> circuitSummary = new List<string>();
        for (int index = 0; index < slots.Count; index++)
        {
          AddCircuitLabelsToReceptacles(
            db,
            ed,
            slotReceptacles[slots[index]].ToArray(),
            scale.PaperInchesPerModelFoot,
            panelName,
            allocations[index].CircuitNumber.ToString()
          );
          circuitSummary.Add(
            $"{allocations[index].CircuitNumber} ({requests[index].ConnectedWatts / 1000.0:0.00} kVA)"
          );
        }

        ed.WriteMessage(
          $"\nRoom circuiting complete for {roomName}: {CountSlotReceptacles(slotReceptacles)} receptacle(s) "
            + $"on {slots.Count} circuit(s): {string.Join(", ", circuitSummary)}."
            + (
              allocations.Count > 0 && allocations[0].RemainingCounts != null
                ? $"\n{FormatPanelCircuitStatus(panelName, allocations[0].RemainingCounts)}"
                : string.Empty
            )
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage(
          $"\nThe room was placed, but circuiting failed: {ex.Message} "
            + "Select the receptacles and run RC to circuit them."
        );
      }
    }

    private static int CountSlotReceptacles(List<ObjectId>[] slotReceptacles)
    {
      int count = 0;
      foreach (List<ObjectId> receptacles in slotReceptacles)
      {
        count += receptacles.Count;
      }
      return count;
    }

    private static Point3d CalculateReceptacleCenter(Database db, List<ObjectId> receptacleIds)
    {
      List<ReceptacleLoadItem> items = ReadReceptacleLoadItems(db, receptacleIds.ToArray(), out _, out _, out _);
      return CalculateReceptacleGroupCenter(items);
    }

    // The AREALABEL room polyline around the point, so the room name does not need typing.
    private static string TryFindTaggedRoomName(Database db, Point3d point)
    {
      List<TaggedRoomBoundary> rooms = new List<TaggedRoomBoundary>();
      using (Transaction transaction = db.TransactionManager.StartOpenCloseTransaction())
      {
        BlockTableRecord space = transaction.GetObject(db.CurrentSpaceId, OpenMode.ForRead, false) as BlockTableRecord;
        if (space == null)
        {
          return string.Empty;
        }

        foreach (ObjectId id in space)
        {
          Autodesk.AutoCAD.DatabaseServices.Polyline polyline =
            transaction.GetObject(id, OpenMode.ForRead, false)
            as Autodesk.AutoCAD.DatabaseServices.Polyline;
          if (
            polyline == null
            || polyline.NumberOfVertices < 3
            || !RoomBoundaryMetadataStore.TryRead(polyline, transaction, out var metadata)
          )
          {
            continue;
          }

          rooms.Add(
            new TaggedRoomBoundary
            {
              ObjectId = id,
              RoomName = metadata.Name,
              BasePoint = metadata.BasePoint,
              RelativeBoundary = BuildRelativeRoomBoundary(polyline, metadata.BasePoint),
            }
          );
        }
      }
      return FindTaggedRoomAtPoint(rooms, point)?.RoomName ?? string.Empty;
    }

    private static bool TryPromptTemplateRoomName(Editor ed, string detectedRoom, out string roomName)
    {
      if (string.IsNullOrWhiteSpace(detectedRoom))
      {
        return TryPromptReceptacleRoomName(ed, out roomName);
      }

      roomName = string.Empty;
      PromptStringOptions options = new PromptStringOptions(
        $"\nEnter room name and number <{detectedRoom}>: "
      )
      {
        AllowSpaces = true,
      };
      PromptResult result = ed.GetString(options);
      if (result.Status != PromptStatus.OK && result.Status != PromptStatus.None)
      {
        return false;
      }

      string typed = result.StringResult ?? string.Empty;
      roomName = Regex.Replace(
        string.IsNullOrWhiteSpace(typed) ? detectedRoom : typed,
        @"\s+",
        " "
      ).Trim().ToUpperInvariant();
      if (roomName.StartsWith("RECEPTACLES - ", StringComparison.OrdinalIgnoreCase))
      {
        roomName = roomName.Substring("RECEPTACLES - ".Length).Trim();
      }
      return roomName.Length > 0;
    }
  }
}
