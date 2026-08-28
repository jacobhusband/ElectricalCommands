using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string ReceptBlockName = "RECEPT";
    private const string AlternateReceptBlockName = "Receptacles";
    private const string ReceptNoteLayerName = "EL-NOTES";
    private const short ReceptNoteLayerColor = 4;
    private const double ReceptModelInchesPerFoot = 12.0;
    private const double ReceptNoteGapInPaperInches = 0.0625;

    [CommandMethod("R", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void InsertReceptacle()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      Database db = doc.Database;
      Editor ed = doc.Editor;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage(
          "\nReceptacle insertion requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
        return;
      }

      if (!ElectricalDrawingSettingsStore.TryReadPanelName(
        db,
        out string panelName))
      {
        ed.WriteMessage(
          "\nReceptacle insertion requires a panel name. " +
          "Run SETPANELNAME (SPN) first.");
        return;
      }

      ObjectId[] impliedReceptacles = GetImpliedReceptacles(
        db,
        ed,
        out int impliedRejectedCount);
      if (impliedReceptacles.Length > 0)
      {
        if (impliedRejectedCount > 0)
        {
          ed.WriteMessage(
            $"\nIgnored {impliedRejectedCount} preselected object(s) that " +
            $"were not {ReceptBlockName} or {AlternateReceptBlockName} blocks.");
        }

        AutomaticallyCircuitReceptacles(
          db,
          ed,
          impliedReceptacles,
          scale.PaperInchesPerModelFoot,
          panelName);
        return;
      }

      PromptPointOptions basePointOptions = new PromptPointOptions(
        "\nSpecify insertion point for receptacle or [Existing]: ")
      {
        AllowNone = false
      };
      basePointOptions.Keywords.Add("Existing");

      PromptPointResult basePointResult = ed.GetPoint(basePointOptions);
      if (basePointResult.Status == PromptStatus.Keyword &&
          string.Equals(
            basePointResult.StringResult,
            "Existing",
            StringComparison.OrdinalIgnoreCase))
      {
        ObjectId[] selectedReceptacles = PromptForExistingReceptacles(
          db,
          ed,
          out int rejectedCount);
        if (selectedReceptacles.Length == 0)
        {
          return;
        }

        if (rejectedCount > 0)
        {
          ed.WriteMessage(
            $"\nIgnored {rejectedCount} selected object(s) that were not " +
            $"{ReceptBlockName} or {AlternateReceptBlockName} blocks.");
        }

        AutomaticallyCircuitReceptacles(
          db,
          ed,
          selectedReceptacles,
          scale.PaperInchesPerModelFoot,
          panelName);
        return;
      }

      if (basePointResult.Status != PromptStatus.OK)
      {
        return;
      }

      Point3d basePointUcs = basePointResult.Value;

      if (!TryPromptReceptacleCircuitNumber(
        ed,
        panelName,
        out string circuitNumber))
      {
        return;
      }

      PromptPointOptions orientOptions = new PromptPointOptions(
        "\nSpecify orientation for receptacle: ")
      {
        BasePoint = basePointUcs,
        UseBasePoint = true,
        AllowNone = false
      };

      PromptPointResult orientResult = ed.GetPoint(orientOptions);
      if (orientResult.Status != PromptStatus.OK)
      {
        return;
      }

      Point3d orientPointUcs = orientResult.Value;
      Vector3d ucsDir = orientPointUcs - basePointUcs;

      if (ucsDir.Length < 1e-6)
      {
        ed.WriteMessage("\nOrientation point cannot be identical to insertion point.");
        return;
      }

      Matrix3d ucsToWcs = ed.CurrentUserCoordinateSystem;
      Point3d insertionPoint = basePointUcs.TransformBy(ucsToWcs);
      Vector3d wcsDir = ucsDir.TransformBy(ucsToWcs);

      double wcsAngle = Math.Atan2(wcsDir.Y, wcsDir.X);
      double blockRotation = wcsAngle - (Math.PI / 2.0);

      ResolveReceptNoteOrientation(
        wcsAngle,
        out double noteRotationDegrees,
        out AttachmentPoint noteAttachment);

      string panelCircuitLabel =
        BuildPanelLabel(panelName) + circuitNumber;
      double blockScale = ResolveReceptBlockScale(
        scale.PaperInchesPerModelFoot);
      double textHeight = ResolveHomerunSymbolSize(
        scale.PaperInchesPerModelFoot);
      double noteGap = ReceptNoteGapInPaperInches * blockScale;

      try
      {
        ObjectId noteLayerId = EnsureReceptNoteLayer(db);
        ObjectId textStyleId = EnsureHomerunTextStyle(db);
        string insertedBlockName = ReceptBlockName;

        using (Transaction transaction = db.TransactionManager.StartTransaction())
        {
          BlockTable blockTable = (BlockTable)transaction.GetObject(
            db.BlockTableId,
            OpenMode.ForRead);

          if (!TryResolveReceptBlockDefinition(
            blockTable,
            out ObjectId blockDefinitionId,
            out insertedBlockName))
          {
            ed.WriteMessage(
              "\nCannot insert receptacle: neither block " +
              $"\"{ReceptBlockName}\" nor \"{AlternateReceptBlockName}\" " +
              "is defined in the current drawing.");
            return;
          }

          BlockTableRecord currentSpace = (BlockTableRecord)transaction.GetObject(
            db.CurrentSpaceId,
            OpenMode.ForWrite);

          BlockReference blockReference = new BlockReference(
            insertionPoint,
            blockDefinitionId);
          blockReference.SetDatabaseDefaults(db);
          blockReference.ScaleFactors = new Scale3d(blockScale);
          blockReference.Rotation = blockRotation;

          currentSpace.AppendEntity(blockReference);
          transaction.AddNewlyCreatedDBObject(blockReference, true);

          AddDefaultAttributes(
            transaction,
            blockDefinitionId,
            blockReference);

          Vector3d noteDirection = Vector3d.YAxis
            .RotateBy(blockRotation, Vector3d.ZAxis);
          Point3d noteLocation = ResolveReceptNoteLocation(
            blockReference,
            insertionPoint,
            noteDirection,
            noteGap);

          MText panelLabel = CreateReceptPanelLabel(
            db,
            noteLocation,
            noteRotationDegrees * Math.PI / 180.0,
            textHeight,
            EscapeMTextPlainText(panelCircuitLabel),
            noteLayerId,
            textStyleId,
            noteAttachment);
          currentSpace.AppendEntity(panelLabel);
          transaction.AddNewlyCreatedDBObject(panelLabel, true);

          transaction.Commit();
        }

        ed.WriteMessage(
          $"\nInserted {insertedBlockName} at " +
          $"{scale.DisplayText} (X/Y/Z scale {FormatNumber(blockScale)}) " +
          $"with circuit label {panelCircuitLabel}.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage(
          $"\nUnable to insert {ReceptBlockName}: {ex.Message}");
      }
    }

    private static bool TryPromptReceptacleCircuitNumber(
      Editor editor,
      string panelName,
      out string circuitNumber)
    {
      circuitNumber = string.Empty;
      PromptStringOptions circuitOptions = new PromptStringOptions(
        $"\nEnter circuit number for panel {panelName} " +
        "(for example 19 or 1/5): ")
      {
        AllowSpaces = false
      };
      PromptResult circuitResult = editor.GetString(circuitOptions);
      if (circuitResult.Status != PromptStatus.OK)
      {
        editor.WriteMessage("\nReceptacle insertion canceled.");
        return false;
      }

      circuitNumber =
        (circuitResult.StringResult ?? string.Empty).Trim().TrimStart('-');
      if (circuitNumber.Length == 0)
      {
        editor.WriteMessage("\nCircuit number cannot be blank.");
        return false;
      }
      return true;
    }

    private static ObjectId[] GetImpliedReceptacles(
      Database database,
      Editor editor,
      out int rejectedCount)
    {
      PromptSelectionResult impliedSelection = editor.SelectImplied();
      if (impliedSelection.Status != PromptStatus.OK ||
          impliedSelection.Value == null)
      {
        rejectedCount = 0;
        return new ObjectId[0];
      }

      return FilterReceptacleIds(
        database,
        impliedSelection.Value.GetObjectIds(),
        out rejectedCount);
    }

    private static ObjectId[] PromptForExistingReceptacles(
      Database database,
      Editor editor,
      out int rejectedCount)
    {
      PromptSelectionOptions selectionOptions = new PromptSelectionOptions
      {
        MessageForAdding =
          $"\nSelect {ReceptBlockName} or {AlternateReceptBlockName} blocks: ",
        AllowDuplicates = false
      };
      SelectionFilter blockFilter = new SelectionFilter(
        new[]
        {
          new TypedValue((int)DxfCode.Start, "INSERT")
        });

      PromptSelectionResult selectionResult = editor.GetSelection(
        selectionOptions,
        blockFilter);
      if (selectionResult.Status != PromptStatus.OK ||
          selectionResult.Value == null)
      {
        rejectedCount = 0;
        editor.WriteMessage("\nReceptacle circuiting canceled.");
        return new ObjectId[0];
      }

      ObjectId[] receptacleIds = FilterReceptacleIds(
        database,
        selectionResult.Value.GetObjectIds(),
        out rejectedCount);
      if (receptacleIds.Length == 0)
      {
        editor.WriteMessage(
          $"\nNo {ReceptBlockName} or {AlternateReceptBlockName} blocks " +
          "were selected.");
      }

      return receptacleIds;
    }

    private static ObjectId[] FilterReceptacleIds(
      Database database,
      ObjectId[] candidateIds,
      out int rejectedCount)
    {
      List<ObjectId> receptacleIds = new List<ObjectId>();
      int candidateCount = candidateIds == null ? 0 : candidateIds.Length;

      if (candidateCount == 0)
      {
        rejectedCount = 0;
        return receptacleIds.ToArray();
      }

      using (Transaction transaction =
        database.TransactionManager.StartTransaction())
      {
        foreach (ObjectId candidateId in candidateIds)
        {
          try
          {
            BlockReference blockReference = transaction.GetObject(
              candidateId,
              OpenMode.ForRead,
              false) as BlockReference;

            if (blockReference != null &&
                blockReference.OwnerId == database.CurrentSpaceId &&
                IsSupportedReceptacleBlock(transaction, blockReference))
            {
              receptacleIds.Add(candidateId);
            }
          }
          catch
          {
            // Ignore stale or otherwise unreadable selection entries.
          }
        }

        transaction.Commit();
      }

      rejectedCount = candidateCount - receptacleIds.Count;
      return receptacleIds.ToArray();
    }

    private static bool IsSupportedReceptacleBlock(
      Transaction transaction,
      BlockReference blockReference)
    {
      ObjectId definitionId = blockReference.IsDynamicBlock
        ? blockReference.DynamicBlockTableRecord
        : blockReference.BlockTableRecord;
      BlockTableRecord definition = transaction.GetObject(
        definitionId,
        OpenMode.ForRead) as BlockTableRecord;

      return definition != null && IsSupportedReceptacleName(definition.Name);
    }

    private static bool IsSupportedReceptacleName(string blockName)
    {
      return string.Equals(
          blockName,
          ReceptBlockName,
          StringComparison.OrdinalIgnoreCase) ||
        string.Equals(
          blockName,
          AlternateReceptBlockName,
          StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryResolveReceptBlockDefinition(
      BlockTable blockTable,
      out ObjectId blockDefinitionId,
      out string resolvedBlockName)
    {
      if (blockTable.Has(ReceptBlockName))
      {
        blockDefinitionId = blockTable[ReceptBlockName];
        resolvedBlockName = ReceptBlockName;
        return true;
      }

      if (blockTable.Has(AlternateReceptBlockName))
      {
        blockDefinitionId = blockTable[AlternateReceptBlockName];
        resolvedBlockName = AlternateReceptBlockName;
        return true;
      }

      blockDefinitionId = ObjectId.Null;
      resolvedBlockName = string.Empty;
      return false;
    }

    private static void AutomaticallyCircuitReceptacles(
      Database database,
      Editor editor,
      ObjectId[] receptacleIds,
      double paperInchesPerModelFoot,
      string panelName)
    {
      if (!ElectricalDrawingSettingsStore.TryReadPanelSchedule(
        database,
        out var panelSchedule))
      {
        editor.WriteMessage(
          "\nAutomatic receptacle circuiting requires a linked panel " +
          "schedule. Run SETPANELSCHEDULE (SPS) first.");
        return;
      }

      CalculateSelectedReceptacleLoad(
        database,
        receptacleIds,
        out double connectedWatts,
        out int duplexCount,
        out int quadCount,
        out int defaultedCount);
      if (connectedWatts <= 0.0)
      {
        editor.WriteMessage(
          "\nNo supported receptacles were available for circuiting.");
        return;
      }

      try
      {
        PanelScheduleAllocationResult allocation =
          PanelScheduleWorkbookAllocator.AllocateReceptacleCircuit(
            panelSchedule.WorkbookPath,
            panelName,
            panelSchedule.CircuitCapacity,
            connectedWatts);

        AddCircuitLabelsToReceptacles(
          database,
          editor,
          receptacleIds,
          paperInchesPerModelFoot,
          panelName,
          allocation.CircuitNumber.ToString());

        editor.WriteMessage(
          $"\nUpdated worksheet \"{allocation.WorksheetName}\", circuit " +
          $"{allocation.CircuitNumber}: {connectedWatts:0} VA " +
          $"({duplexCount} duplex, {quadCount} quad" +
          (defaultedCount > 0
            ? $", {defaultedCount} defaulted to 180 VA"
            : string.Empty) +
          ").");
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to automatically circuit the receptacles: {ex.Message}");
      }
    }

    private static void CalculateSelectedReceptacleLoad(
      Database database,
      ObjectId[] receptacleIds,
      out double connectedWatts,
      out int duplexCount,
      out int quadCount,
      out int defaultedCount)
    {
      connectedWatts = 0.0;
      duplexCount = 0;
      quadCount = 0;
      defaultedCount = 0;

      using (Transaction transaction =
        database.TransactionManager.StartOpenCloseTransaction())
      {
        foreach (ObjectId receptacleId in receptacleIds)
        {
          BlockReference blockReference = transaction.GetObject(
            receptacleId,
            OpenMode.ForRead,
            false) as BlockReference;
          if (blockReference == null ||
              !IsSupportedReceptacleBlock(transaction, blockReference))
          {
            continue;
          }

          string visibilityState = ResolveReceptacleVisibilityState(
            blockReference);
          if (visibilityState.StartsWith(
            "QUAD",
            StringComparison.OrdinalIgnoreCase))
          {
            connectedWatts += 360.0;
            quadCount++;
          }
          else
          {
            connectedWatts += 180.0;
            if (visibilityState.StartsWith(
              "DUPLEX",
              StringComparison.OrdinalIgnoreCase))
            {
              duplexCount++;
            }
            else
            {
              defaultedCount++;
            }
          }
        }
      }
    }

    private static string ResolveReceptacleVisibilityState(
      BlockReference blockReference)
    {
      if (!blockReference.IsDynamicBlock)
      {
        return string.Empty;
      }

      try
      {
        DynamicBlockReferencePropertyCollection properties =
          blockReference.DynamicBlockReferencePropertyCollection;
        foreach (DynamicBlockReferenceProperty property in properties)
        {
          string propertyName = property.PropertyName ?? string.Empty;
          string value = Convert.ToString(property.Value) ?? string.Empty;
          if (propertyName.IndexOf(
                "VISIBILITY",
                StringComparison.OrdinalIgnoreCase) >= 0 ||
              value.StartsWith("DUPLEX", StringComparison.OrdinalIgnoreCase) ||
              value.StartsWith("QUAD", StringComparison.OrdinalIgnoreCase))
          {
            return value.Trim();
          }
        }
      }
      catch
      {
        // A nonstandard or unreadable dynamic block defaults to 180 VA.
      }

      return string.Empty;
    }

    private static void AddCircuitLabelsToReceptacles(
      Database database,
      Editor editor,
      ObjectId[] receptacleIds,
      double paperInchesPerModelFoot,
      string panelName,
      string circuitNumber)
    {
      string panelCircuitLabel = BuildPanelLabel(panelName) + circuitNumber;
      double blockScale = ResolveReceptBlockScale(paperInchesPerModelFoot);
      double textHeight = ResolveHomerunSymbolSize(paperInchesPerModelFoot);
      double noteGap = ReceptNoteGapInPaperInches * blockScale;

      try
      {
        ObjectId noteLayerId = EnsureReceptNoteLayer(database);
        ObjectId textStyleId = EnsureHomerunTextStyle(database);
        int labeledCount = 0;
        int skippedCount = 0;

        using (Transaction transaction =
          database.TransactionManager.StartTransaction())
        {
          BlockTableRecord currentSpace =
            (BlockTableRecord)transaction.GetObject(
              database.CurrentSpaceId,
              OpenMode.ForWrite);

          foreach (ObjectId receptacleId in receptacleIds)
          {
            BlockReference blockReference = transaction.GetObject(
              receptacleId,
              OpenMode.ForRead,
              false) as BlockReference;
            if (blockReference == null ||
                blockReference.OwnerId != database.CurrentSpaceId ||
                !IsSupportedReceptacleBlock(transaction, blockReference))
            {
              skippedCount++;
              continue;
            }

            Vector3d facingDirection = Vector3d.YAxis.TransformBy(
              blockReference.BlockTransform);
            double planarLength = Math.Sqrt(
              facingDirection.X * facingDirection.X +
              facingDirection.Y * facingDirection.Y);
            if (planarLength < 1e-9)
            {
              skippedCount++;
              continue;
            }

            Vector3d noteDirection = new Vector3d(
              facingDirection.X / planarLength,
              facingDirection.Y / planarLength,
              0.0);
            double facingAngle = Math.Atan2(
              noteDirection.Y,
              noteDirection.X);
            ResolveReceptNoteOrientation(
              facingAngle,
              out double noteRotationDegrees,
              out AttachmentPoint noteAttachment);

            Point3d noteLocation = ResolveReceptNoteLocation(
              blockReference,
              blockReference.Position,
              noteDirection,
              noteGap);
            MText panelLabel = CreateReceptPanelLabel(
              database,
              noteLocation,
              noteRotationDegrees * Math.PI / 180.0,
              textHeight,
              EscapeMTextPlainText(panelCircuitLabel),
              noteLayerId,
              textStyleId,
              noteAttachment);
            currentSpace.AppendEntity(panelLabel);
            transaction.AddNewlyCreatedDBObject(panelLabel, true);
            labeledCount++;
          }

          transaction.Commit();
        }

        editor.WriteMessage(
          $"\nAdded circuit label {panelCircuitLabel} to " +
          $"{labeledCount} receptacle(s)." +
          (skippedCount > 0
            ? $" Skipped {skippedCount} receptacle(s)."
            : string.Empty));
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to circuit selected receptacles: {ex.Message}");
      }
    }

    private static void ResolveReceptNoteOrientation(
      double facingAngle,
      out double noteRotationDegrees,
      out AttachmentPoint noteAttachment)
    {
      double degrees = (facingAngle * 180.0 / Math.PI) % 360.0;
      if (degrees < 0)
      {
        degrees += 360.0;
      }

      if (degrees >= 45.0 && degrees < 135.0)
      {
        // Facing North (+Y)
        noteRotationDegrees = 90.0;
        noteAttachment = AttachmentPoint.MiddleLeft;
      }
      else if (degrees >= 135.0 && degrees < 225.0)
      {
        // Facing West (-X)
        noteRotationDegrees = 0.0;
        noteAttachment = AttachmentPoint.MiddleRight;
      }
      else if (degrees >= 225.0 && degrees < 315.0)
      {
        // Facing South (-Y)
        noteRotationDegrees = 270.0;
        noteAttachment = AttachmentPoint.MiddleLeft;
      }
      else
      {
        // Facing East (+X)
        noteRotationDegrees = 0.0;
        noteAttachment = AttachmentPoint.MiddleLeft;
      }
    }

    private static double ResolveReceptBlockScale(
      double paperInchesPerModelFoot)
    {
      return ReceptModelInchesPerFoot / paperInchesPerModelFoot;
    }

    private static ObjectId EnsureReceptNoteLayer(Database database)
    {
      using (Transaction transaction =
        database.TransactionManager.StartTransaction())
      {
        LayerTable layers = (LayerTable)transaction.GetObject(
          database.LayerTableId,
          OpenMode.ForRead);

        if (layers.Has(ReceptNoteLayerName))
        {
          ObjectId existingId = layers[ReceptNoteLayerName];
          transaction.Commit();
          return existingId;
        }

        layers.UpgradeOpen();
        LayerTableRecord layer = new LayerTableRecord
        {
          Name = ReceptNoteLayerName,
          Color = Color.FromColorIndex(
            ColorMethod.ByAci,
            ReceptNoteLayerColor)
        };
        ObjectId layerId = layers.Add(layer);
        transaction.AddNewlyCreatedDBObject(layer, true);
        transaction.Commit();
        return layerId;
      }
    }

    private static Point3d ResolveReceptNoteLocation(
      BlockReference blockReference,
      Point3d insertionPoint,
      Vector3d noteDirection,
      double gap)
    {
      try
      {
        Extents3d extents = blockReference.GeometricExtents;
        Point3d center = new Point3d(
          (extents.MinPoint.X + extents.MaxPoint.X) / 2.0,
          (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0,
          insertionPoint.Z);
        double halfWidth =
          (extents.MaxPoint.X - extents.MinPoint.X) / 2.0;
        double halfHeight =
          (extents.MaxPoint.Y - extents.MinPoint.Y) / 2.0;
        double distanceToEdge =
          Math.Abs(noteDirection.X) * halfWidth +
          Math.Abs(noteDirection.Y) * halfHeight;

        return center + noteDirection * (distanceToEdge + gap);
      }
      catch
      {
        return insertionPoint + noteDirection * gap;
      }
    }

    private static MText CreateReceptPanelLabel(
      Database database,
      Point3d location,
      double rotation,
      double textHeight,
      string contents,
      ObjectId layerId,
      ObjectId textStyleId,
      AttachmentPoint attachment)
    {
      MText text = new MText();
      text.SetDatabaseDefaults(database);
      text.LayerId = layerId;
      text.ColorIndex = 256;
      text.Location = location;
      text.Contents = contents;
      text.Attachment = attachment;
      text.TextHeight = textHeight;
      text.Width = 0.0;
      text.Rotation = rotation;
      text.Annotative = AnnotativeStates.False;
      text.TextStyleId = textStyleId;
      return text;
    }

    private static void AddDefaultAttributes(
      Transaction transaction,
      ObjectId blockDefinitionId,
      BlockReference blockReference)
    {
      BlockTableRecord blockDefinition =
        (BlockTableRecord)transaction.GetObject(
          blockDefinitionId,
          OpenMode.ForRead);

      foreach (ObjectId entityId in blockDefinition)
      {
        AttributeDefinition attributeDefinition =
          transaction.GetObject(
            entityId,
            OpenMode.ForRead) as AttributeDefinition;

        if (attributeDefinition == null || attributeDefinition.Constant)
        {
          continue;
        }

        AttributeReference attributeReference = new AttributeReference();
        attributeReference.SetAttributeFromBlock(
          attributeDefinition,
          blockReference.BlockTransform);
        attributeReference.TextString = attributeDefinition.TextString;

        blockReference.AttributeCollection.AppendAttribute(attributeReference);
        transaction.AddNewlyCreatedDBObject(attributeReference, true);
      }
    }
  }
}
