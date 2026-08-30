using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
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
    internal const double ReceptacleLoadUnitKva = 0.18;
    internal const int DefaultReceptacleCircuitLoadUnits = 7;
    internal const int MaximumReceptacleCircuitLoadUnits = 13;
    internal const double DefaultReceptacleCircuitMaxKva =
      DefaultReceptacleCircuitLoadUnits * ReceptacleLoadUnitKva;
    internal const double MaximumReceptacleCircuitLoadKva =
      MaximumReceptacleCircuitLoadUnits * ReceptacleLoadUnitKva;
    private const double ReceptNoteGapInPaperInches = 0.0625;
    private const double ReceptNoteClearanceInPaperInches = 0.03125;
    private const double ReceptNoteSearchStepInPaperInches = 0.0625;
    private const double ReceptNoteMaximumGapInPaperInches = 0.25;

    [CommandMethod("R", CommandFlags.Modal)]
    public static void PlaceReceptacles()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      Database db = doc.Database;
      Editor ed = doc.Editor;

      ElectricalDrawingSettingsStore.ScaleSetting scale;
      if (TrySetReceptacleScaleFromActiveViewport(
        db,
        ed,
        out scale,
        out bool isEditingViewport,
        out string viewportScaleError))
      {
        ed.WriteMessage(
          $"\nDrawing scale automatically set to {scale.DisplayText} " +
          "from the active viewport.");
        HomerunSettingsPalette.Refresh();
        HomerunSettingsPalette.SetStatus(
          $"Scale set to {scale.DisplayText} from the active viewport.");
      }
      else if (isEditingViewport)
      {
        ed.WriteMessage(
          "\nReceptacle placement could not determine the active " +
          $"viewport scale: {viewportScaleError}");
        return;
      }
      else if (!ElectricalDrawingSettingsStore.TryReadScale(db, out scale))
      {
        ed.WriteMessage(
          "\nReceptacle placement requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
        return;
      }

      if (!TryReadReceptaclePlacementOptions(
        db,
        out ObjectId blockDefinitionId,
        out string blockName,
        out List<string> visibilityStates,
        out string selectedVisibilityState,
        out string placementOptionsError))
      {
        ed.WriteMessage($"\nCannot start receptacle placement: {placementOptionsError}");
        return;
      }

      double blockScale = ResolveReceptBlockScale(
        scale.PaperInchesPerModelFoot);
      int placedCount = 0;

      if (visibilityStates.Count == 0)
      {
        ed.WriteMessage(
          $"\n{blockName} has no selectable receptacle visibility states. " +
          "The block's default appearance will be used.");
      }
      else
      {
        ed.WriteMessage(
          $"\nReceptacle type set to {selectedVisibilityState}. " +
          "Press S at the insertion-point prompt to change it.");
      }

      while (true)
      {
        string typeDescription = string.IsNullOrWhiteSpace(
          selectedVisibilityState)
          ? blockName
          : selectedVisibilityState;
        string finishDescription = placedCount > 0
          ? " <Enter to finish>"
          : string.Empty;
        PromptPointOptions insertionOptions = new PromptPointOptions(
          $"\nSpecify insertion point for {typeDescription}" +
          (visibilityStates.Count > 0 ? " or [Set]" : string.Empty) +
          $":{finishDescription} ")
        {
          AllowNone = placedCount > 0
        };
        if (visibilityStates.Count > 0)
        {
          insertionOptions.Keywords.Add("Set");
        }

        PromptPointResult insertionResult = ed.GetPoint(insertionOptions);
        if (insertionResult.Status == PromptStatus.Keyword &&
            string.Equals(
              insertionResult.StringResult,
              "Set",
              StringComparison.OrdinalIgnoreCase))
        {
          if (!TryPromptReceptacleVisibilityState(
            ed,
            visibilityStates,
            selectedVisibilityState,
            out selectedVisibilityState))
          {
            return;
          }
          continue;
        }

        if (insertionResult.Status != PromptStatus.OK)
        {
          return;
        }

        Point3d insertionPointUcs = insertionResult.Value;
        PromptPointOptions orientationOptions = new PromptPointOptions(
          "\nSpecify orientation for receptacle: ")
        {
          BasePoint = insertionPointUcs,
          UseBasePoint = true,
          AllowNone = false
        };
        PromptPointResult orientationResult = ed.GetPoint(
          orientationOptions);
        if (orientationResult.Status != PromptStatus.OK)
        {
          return;
        }

        Vector3d ucsDirection =
          orientationResult.Value - insertionPointUcs;
        if (ucsDirection.Length < 1e-6)
        {
          ed.WriteMessage(
            "\nOrientation point cannot be identical to insertion point.");
          continue;
        }

        Matrix3d ucsToWcs = ed.CurrentUserCoordinateSystem;
        Point3d insertionPoint = insertionPointUcs.TransformBy(ucsToWcs);
        Vector3d wcsDirection = ucsDirection.TransformBy(ucsToWcs);
        double blockRotation =
          Math.Atan2(wcsDirection.Y, wcsDirection.X) - (Math.PI / 2.0);

        try
        {
          InsertReceptacleBlock(
            db,
            blockDefinitionId,
            insertionPoint,
            blockRotation,
            blockScale,
            selectedVisibilityState);
          placedCount++;
          ed.WriteMessage(
            $"\nPlaced {typeDescription} receptacle " +
            $"at {scale.DisplayText}. Continue placing or press Enter to finish.");
        }
        catch (System.Exception ex)
        {
          ed.WriteMessage(
            $"\nUnable to insert {blockName}: {ex.Message}");
          return;
        }
      }
    }

    private static bool TrySetReceptacleScaleFromActiveViewport(
      Database database,
      Editor editor,
      out ElectricalDrawingSettingsStore.ScaleSetting scale,
      out bool isEditingViewport,
      out string errorMessage)
    {
      scale = null;
      isEditingViewport = false;
      errorMessage = string.Empty;

      if (database.TileMode)
      {
        return false;
      }

      try
      {
        isEditingViewport =
          Convert.ToInt16(Application.GetSystemVariable("CVPORT")) > 1;
      }
      catch (System.Exception ex)
      {
        errorMessage = ex.Message;
        return false;
      }

      if (!isEditingViewport)
      {
        return false;
      }

      double scaleDenominator = ResolveViewportScaleDenominator(
        editor,
        database);
      if (scaleDenominator <= 0.0 ||
          double.IsNaN(scaleDenominator) ||
          double.IsInfinity(scaleDenominator))
      {
        errorMessage = "AutoCAD did not return a valid viewport scale.";
        return false;
      }

      double paperInchesPerModelFoot =
        ReceptModelInchesPerFoot / scaleDenominator;
      string displayText = FormatArchitecturalScale(
        paperInchesPerModelFoot);

      try
      {
        ElectricalDrawingSettingsStore.WriteScale(
          database,
          paperInchesPerModelFoot,
          displayText);
        scale = new ElectricalDrawingSettingsStore.ScaleSetting
        {
          PaperInchesPerModelFoot = paperInchesPerModelFoot,
          DisplayText = displayText,
        };
        return true;
      }
      catch (System.Exception ex)
      {
        errorMessage = ex.Message;
        return false;
      }
    }

    [CommandMethod("RC", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void InsertReceptacle()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      Database db = doc.Database;
      Editor ed = doc.Editor;

      if (!ElectricalDrawingSettingsStore.TryReadPanelName(
        db,
        out string panelName))
      {
        ed.WriteMessage(
          "\nReceptacle insertion requires a panel name. " +
          "Run SETPANELNAME (SPN) first.");
        return;
      }

      bool hasUsablePanelSchedule =
        ElectricalDrawingSettingsStore.TryReadPanelSchedule(
          db,
          out var panelSchedule) &&
        File.Exists(panelSchedule.WorkbookPath);
      if (!hasUsablePanelSchedule)
      {
        ed.WriteMessage(
          "\nRC requires a linked panel schedule. Select the panel " +
          "schedule now before receptacle circuiting begins.");
        SetPanelScheduleCommand();

        hasUsablePanelSchedule =
          ElectricalDrawingSettingsStore.TryReadPanelSchedule(
            db,
            out panelSchedule) &&
          File.Exists(panelSchedule.WorkbookPath);
        if (!hasUsablePanelSchedule)
        {
          ed.WriteMessage(
            "\nRC canceled because no usable panel schedule was linked.");
          return;
        }
      }

      ElectricalDrawingSettingsStore.ScaleSetting scale;
      if (TrySetReceptacleScaleFromActiveViewport(
        db,
        ed,
        out scale,
        out bool isEditingViewport,
        out string viewportScaleError))
      {
        ed.WriteMessage(
          $"\nDrawing scale automatically set to {scale.DisplayText} " +
          "from the active viewport.");
        HomerunSettingsPalette.Refresh();
        HomerunSettingsPalette.SetStatus(
          $"Scale set to {scale.DisplayText} from the active viewport.");
      }
      else if (isEditingViewport)
      {
        ed.WriteMessage(
          "\nReceptacle insertion could not determine the active " +
          $"viewport scale: {viewportScaleError}");
        return;
      }
      else if (!ElectricalDrawingSettingsStore.TryReadScale(db, out scale))
      {
        ed.WriteMessage(
          "\nReceptacle insertion requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
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
        "\nSpecify insertion point for receptacle or " +
        "[Existing/Room/Dedicated]: ")
      {
        AllowNone = false
      };
      basePointOptions.Keywords.Add("Existing");
      basePointOptions.Keywords.Add("Room");
      basePointOptions.Keywords.Add("Dedicated");

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

      if (basePointResult.Status == PromptStatus.Keyword &&
          string.Equals(
            basePointResult.StringResult,
            "Dedicated",
            StringComparison.OrdinalIgnoreCase))
      {
        if (!TryPromptDedicatedReceptacle(
          db,
          ed,
          out ObjectId dedicatedReceptacleId) ||
            !TryPromptDedicatedEquipment(
              ed,
              out DedicatedEquipmentLoad equipment))
        {
          return;
        }

        AutomaticallyCircuitDedicatedReceptacle(
          db,
          ed,
          dedicatedReceptacleId,
          scale.PaperInchesPerModelFoot,
          panelName,
          equipment);
        return;
      }

      if (basePointResult.Status == PromptStatus.Keyword &&
          string.Equals(
            basePointResult.StringResult,
            "Room",
            StringComparison.OrdinalIgnoreCase))
      {
        double configuredMaximumKva =
          ResolveReceptacleCircuitMaximumKva(db);
        ed.WriteMessage(
          $"\nRoom circuits will be balanced up to " +
          $"{configuredMaximumKva:0.00} kVA each. Change RC Max kVA in " +
          "HRS to use a different project limit.");

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

        if (!TryPromptReceptacleRoomName(ed, out string roomName))
        {
          return;
        }

        AutomaticallyCircuitRoomReceptacles(
          db,
          ed,
          selectedReceptacles,
          scale.PaperInchesPerModelFoot,
          panelName,
          roomName);
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
      double noteClearance =
        ReceptNoteClearanceInPaperInches * blockScale;
      double noteSearchStep = Math.Max(
        ReceptNoteSearchStepInPaperInches * blockScale,
        textHeight * 0.65);
      double noteMaximumGap = Math.Max(
        noteGap,
        ReceptNoteMaximumGapInPaperInches * blockScale);
      bool labelUsedAlternatePosition = false;

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

          labelUsedAlternatePosition = PlaceReceptPanelLabelWithoutOverlap(
            ed,
            transaction,
            blockReference,
            panelLabel,
            noteDirection,
            noteGap,
            noteClearance,
            noteSearchStep,
            noteMaximumGap,
            null);

          transaction.Commit();
        }

        ed.WriteMessage(
          $"\nInserted {insertedBlockName} at " +
          $"{scale.DisplayText} (X/Y/Z scale {FormatNumber(blockScale)}) " +
          $"with circuit label {panelCircuitLabel}." +
          (labelUsedAlternatePosition
            ? " The label was placed in an alternate position to avoid nearby objects."
            : string.Empty));
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage(
          $"\nUnable to insert {ReceptBlockName}: {ex.Message}");
      }
    }

    private static bool TryPromptReceptacleRoomName(
      Editor editor,
      out string roomName)
    {
      roomName = string.Empty;
      PromptStringOptions roomOptions = new PromptStringOptions(
        "\nEnter room name and number " +
        "(for example CONFERENCE ROOM 108): ")
      {
        AllowSpaces = true
      };
      PromptResult roomResult = editor.GetString(roomOptions);
      if (roomResult.Status != PromptStatus.OK)
      {
        editor.WriteMessage("\nRoom receptacle circuiting canceled.");
        return false;
      }

      roomName = Regex.Replace(
        roomResult.StringResult ?? string.Empty,
        @"\s+",
        " ").Trim().ToUpperInvariant();
      if (roomName.StartsWith(
        "RECEPTACLES - ",
        StringComparison.OrdinalIgnoreCase))
      {
        roomName = roomName.Substring("RECEPTACLES - ".Length).Trim();
      }
      if (roomName.Length == 0)
      {
        editor.WriteMessage("\nRoom name cannot be blank.");
        return false;
      }
      return true;
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

    private static bool TryReadReceptaclePlacementOptions(
      Database database,
      out ObjectId blockDefinitionId,
      out string blockName,
      out List<string> visibilityStates,
      out string currentVisibilityState,
      out string errorMessage)
    {
      blockDefinitionId = ObjectId.Null;
      blockName = string.Empty;
      visibilityStates = new List<string>();
      currentVisibilityState = string.Empty;
      errorMessage = string.Empty;

      try
      {
        using (Transaction transaction =
          database.TransactionManager.StartTransaction())
        {
          BlockTable blockTable = (BlockTable)transaction.GetObject(
            database.BlockTableId,
            OpenMode.ForRead);
          if (!TryResolveReceptBlockDefinition(
            blockTable,
            out blockDefinitionId,
            out blockName))
          {
            errorMessage =
              $"neither block \"{ReceptBlockName}\" nor " +
              $"\"{AlternateReceptBlockName}\" is defined in the current drawing.";
            return false;
          }

          BlockTableRecord currentSpace =
            (BlockTableRecord)transaction.GetObject(
              database.CurrentSpaceId,
              OpenMode.ForWrite);
          BlockReference probeReference = new BlockReference(
            Point3d.Origin,
            blockDefinitionId);
          currentSpace.AppendEntity(probeReference);
          transaction.AddNewlyCreatedDBObject(probeReference, true);

          ReadReceptacleVisibilityStates(
            probeReference,
            visibilityStates,
            out currentVisibilityState);

          // The temporary reference exists only so AutoCAD exposes the dynamic
          // block arguments. Aborting leaves the drawing unchanged.
          transaction.Abort();
        }

        if (visibilityStates.Count > 0 &&
            string.IsNullOrWhiteSpace(currentVisibilityState))
        {
          currentVisibilityState = visibilityStates[0];
        }
        return true;
      }
      catch (System.Exception ex)
      {
        errorMessage = ex.Message;
        return false;
      }
    }

    private static void ReadReceptacleVisibilityStates(
      BlockReference blockReference,
      List<string> visibilityStates,
      out string currentVisibilityState)
    {
      currentVisibilityState = string.Empty;
      if (!blockReference.IsDynamicBlock)
      {
        return;
      }

      DynamicBlockReferenceProperty fallbackProperty = null;
      foreach (DynamicBlockReferenceProperty property in
        blockReference.DynamicBlockReferencePropertyCollection)
      {
        object[] allowedValues;
        try
        {
          allowedValues = property.GetAllowedValues();
        }
        catch
        {
          continue;
        }

        if (allowedValues == null || allowedValues.Length == 0)
        {
          continue;
        }

        bool isVisibilityProperty =
          (property.PropertyName ?? string.Empty).IndexOf(
            "VISIBILITY",
            StringComparison.OrdinalIgnoreCase) >= 0;
        bool containsReceptacleStates = false;
        foreach (object allowedValue in allowedValues)
        {
          if (LooksLikeReceptacleVisibilityState(
            Convert.ToString(allowedValue)))
          {
            containsReceptacleStates = true;
            break;
          }
        }

        if (isVisibilityProperty)
        {
          AddReceptacleVisibilityStates(
            visibilityStates,
            allowedValues);
          currentVisibilityState =
            (Convert.ToString(property.Value) ?? string.Empty).Trim();
          return;
        }

        if (fallbackProperty == null && containsReceptacleStates)
        {
          fallbackProperty = property;
        }
      }

      if (fallbackProperty == null)
      {
        return;
      }

      object[] fallbackValues = fallbackProperty.GetAllowedValues();
      AddReceptacleVisibilityStates(visibilityStates, fallbackValues);
      currentVisibilityState =
        (Convert.ToString(fallbackProperty.Value) ?? string.Empty).Trim();
    }

    private static void AddReceptacleVisibilityStates(
      List<string> visibilityStates,
      object[] allowedValues)
    {
      foreach (object allowedValue in allowedValues)
      {
        string state = (Convert.ToString(allowedValue) ?? string.Empty).Trim();
        if (state.Length == 0)
        {
          continue;
        }

        bool alreadyAdded = false;
        foreach (string existingState in visibilityStates)
        {
          if (string.Equals(
            existingState,
            state,
            StringComparison.OrdinalIgnoreCase))
          {
            alreadyAdded = true;
            break;
          }
        }

        if (!alreadyAdded)
        {
          visibilityStates.Add(state);
        }
      }
    }

    private static bool LooksLikeReceptacleVisibilityState(string value)
    {
      string state = (value ?? string.Empty).Trim();
      return state.StartsWith("DUPLEX", StringComparison.OrdinalIgnoreCase) ||
        state.StartsWith("QUAD", StringComparison.OrdinalIgnoreCase) ||
        state.StartsWith("SIMPLEX", StringComparison.OrdinalIgnoreCase) ||
        state.StartsWith("SPECIAL", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryPromptReceptacleVisibilityState(
      Editor editor,
      List<string> visibilityStates,
      string currentVisibilityState,
      out string selectedVisibilityState)
    {
      selectedVisibilityState = currentVisibilityState;
      if (visibilityStates == null || visibilityStates.Count == 0)
      {
        return true;
      }

      PromptKeywordOptions typeOptions = new PromptKeywordOptions(
        $"\nSelect receptacle type <{currentVisibilityState}>: ")
      {
        AllowNone = true
      };
      List<string> keywordNames = new List<string>();
      string defaultKeyword = string.Empty;
      for (int index = 0; index < visibilityStates.Count; index++)
      {
        string state = visibilityStates[index];
        string keywordName = $"ReceptacleType{index + 1}";
        string localName = Regex.IsMatch(
          state,
          @"^[A-Za-z][A-Za-z0-9_-]*$")
          ? state
          : keywordName;
        typeOptions.Keywords.Add(
          keywordName,
          localName,
          state,
          true,
          true);
        keywordNames.Add(keywordName);
        if (string.Equals(
          state,
          currentVisibilityState,
          StringComparison.OrdinalIgnoreCase))
        {
          defaultKeyword = keywordName;
        }
      }

      if (defaultKeyword.Length == 0)
      {
        defaultKeyword = keywordNames[0];
      }
      typeOptions.Keywords.Default = defaultKeyword;

      PromptResult typeResult = editor.GetKeywords(typeOptions);
      if (typeResult.Status == PromptStatus.None)
      {
        return true;
      }
      if (typeResult.Status != PromptStatus.OK)
      {
        return false;
      }

      for (int index = 0; index < visibilityStates.Count; index++)
      {
        if (string.Equals(
            typeResult.StringResult,
            keywordNames[index],
            StringComparison.OrdinalIgnoreCase) ||
          string.Equals(
            typeResult.StringResult,
            visibilityStates[index],
            StringComparison.OrdinalIgnoreCase))
        {
          selectedVisibilityState = visibilityStates[index];
          editor.WriteMessage(
            $"\nReceptacle type set to {selectedVisibilityState}.");
          return true;
        }
      }

      editor.WriteMessage("\nThe selected receptacle type was not recognized.");
      return false;
    }

    private static void InsertReceptacleBlock(
      Database database,
      ObjectId blockDefinitionId,
      Point3d insertionPoint,
      double blockRotation,
      double blockScale,
      string visibilityState)
    {
      using (Transaction transaction =
        database.TransactionManager.StartTransaction())
      {
        BlockTableRecord currentSpace =
          (BlockTableRecord)transaction.GetObject(
            database.CurrentSpaceId,
            OpenMode.ForWrite);
        BlockReference blockReference = new BlockReference(
          insertionPoint,
          blockDefinitionId);
        blockReference.SetDatabaseDefaults(database);
        blockReference.ScaleFactors = new Scale3d(blockScale);
        blockReference.Rotation = blockRotation;

        currentSpace.AppendEntity(blockReference);
        transaction.AddNewlyCreatedDBObject(blockReference, true);

        if (!string.IsNullOrWhiteSpace(visibilityState) &&
            !TrySetReceptacleVisibilityState(
              blockReference,
              visibilityState))
        {
          throw new InvalidOperationException(
            $"The block does not expose receptacle type {visibilityState}.");
        }

        AddDefaultAttributes(
          transaction,
          blockDefinitionId,
          blockReference);
        blockReference.RecordGraphicsModified(true);
        transaction.Commit();
      }
    }

    private static bool TrySetReceptacleVisibilityState(
      BlockReference blockReference,
      string visibilityState)
    {
      if (!blockReference.IsDynamicBlock)
      {
        return false;
      }

      DynamicBlockReferenceProperty fallbackProperty = null;
      object fallbackValue = null;
      foreach (DynamicBlockReferenceProperty property in
        blockReference.DynamicBlockReferencePropertyCollection)
      {
        if (property.ReadOnly)
        {
          continue;
        }

        object matchingValue = null;
        object[] allowedValues;
        try
        {
          allowedValues = property.GetAllowedValues();
        }
        catch
        {
          continue;
        }

        foreach (object allowedValue in allowedValues)
        {
          if (string.Equals(
            Convert.ToString(allowedValue),
            visibilityState,
            StringComparison.OrdinalIgnoreCase))
          {
            matchingValue = allowedValue;
            break;
          }
        }

        if (matchingValue == null)
        {
          continue;
        }

        if ((property.PropertyName ?? string.Empty).IndexOf(
          "VISIBILITY",
          StringComparison.OrdinalIgnoreCase) >= 0)
        {
          property.Value = matchingValue;
          return true;
        }

        if (fallbackProperty == null)
        {
          fallbackProperty = property;
          fallbackValue = matchingValue;
        }
      }

      if (fallbackProperty != null)
      {
        fallbackProperty.Value = fallbackValue;
        return true;
      }
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

    private static void AutomaticallyCircuitRoomReceptacles(
      Database database,
      Editor editor,
      ObjectId[] receptacleIds,
      double paperInchesPerModelFoot,
      string panelName,
      string roomName)
    {
      if (!ElectricalDrawingSettingsStore.TryReadPanelSchedule(
        database,
        out var panelSchedule))
      {
        editor.WriteMessage(
          "\nAutomatic room circuiting requires a linked panel schedule. " +
          "Run SETPANELSCHEDULE (SPS) first.");
        return;
      }

      try
      {
        List<ReceptacleLoadItem> receptacles = ReadReceptacleLoadItems(
          database,
          receptacleIds,
          out int duplexCount,
          out int quadCount,
          out int defaultedCount);
        if (receptacles.Count == 0)
        {
          editor.WriteMessage(
            "\nNo supported receptacles were available for room circuiting.");
          return;
        }

        double maximumCircuitKva =
          ResolveReceptacleCircuitMaximumKva(database);
        int maximumLoadUnits = (int)Math.Round(
          maximumCircuitKva / ReceptacleLoadUnitKva);

        List<ReceptacleCircuitGroup> groups =
          BuildRoomReceptacleCircuitGroups(
            receptacles,
            maximumLoadUnits);
        string loadDescription = "RECEPTACLES - " + roomName;
        List<PanelScheduleCircuitRequest> requests =
          new List<PanelScheduleCircuitRequest>();
        foreach (ReceptacleCircuitGroup group in groups)
        {
          requests.Add(new PanelScheduleCircuitRequest
          {
            ConnectedWatts = group.LoadUnits * 180.0,
            LoadDescription = loadDescription,
          });
        }

        List<PanelScheduleAllocationResult> allocations =
          PanelScheduleWorkbookAllocator.AllocateReceptacleCircuits(
            panelSchedule.WorkbookPath,
            panelName,
            panelSchedule.CircuitCapacity,
            requests);
        if (allocations.Count != groups.Count)
        {
          throw new InvalidOperationException(
            "The panel schedule did not return every requested room circuit.");
        }

        List<string> circuitSummaries = new List<string>();
        double totalWatts = 0.0;
        for (int index = 0; index < groups.Count; index++)
        {
          ReceptacleCircuitGroup group = groups[index];
          PanelScheduleAllocationResult allocation = allocations[index];
          double connectedWatts = group.LoadUnits * 180.0;
          totalWatts += connectedWatts;

          AddCircuitLabelsToReceptacles(
            database,
            editor,
            group.GetObjectIds(),
            paperInchesPerModelFoot,
            panelName,
            allocation.CircuitNumber.ToString());
          circuitSummaries.Add(
            $"{allocation.CircuitNumber} ({connectedWatts / 1000.0:0.00} kVA)");
        }

        editor.WriteMessage(
          $"\nRoom circuiting complete for {roomName}: " +
          $"{receptacles.Count} receptacle(s), {totalWatts / 1000.0:0.00} " +
          $"kVA across {groups.Count} circuit(s): " +
          string.Join(", ", circuitSummaries) + "." +
          (defaultedCount > 0
            ? $" {defaultedCount} nonstandard receptacle block(s) defaulted " +
              $"to {ReceptacleLoadUnitKva:0.00} kVA."
            : string.Empty) +
          $" Loads: {duplexCount} duplex, {quadCount} quad. " +
          $"Maximum circuit load: {maximumCircuitKva:0.00} kVA " +
          "(change in HRS).");
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to circuit receptacles for room {roomName}: {ex.Message}");
      }
    }

    private static double ResolveReceptacleCircuitMaximumKva(
      Database database)
    {
      return ElectricalDrawingSettingsStore.TryReadReceptacleCircuitMaxKva(
          database,
          out double maximumKva) &&
        maximumKva > 0.0
        ? maximumKva
        : DefaultReceptacleCircuitMaxKva;
    }

    private static List<ReceptacleLoadItem> ReadReceptacleLoadItems(
      Database database,
      ObjectId[] receptacleIds,
      out int duplexCount,
      out int quadCount,
      out int defaultedCount)
    {
      List<ReceptacleLoadItem> receptacles =
        new List<ReceptacleLoadItem>();
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
          int loadUnits;
          if (visibilityState.StartsWith(
            "QUAD",
            StringComparison.OrdinalIgnoreCase))
          {
            loadUnits = 2;
            quadCount++;
          }
          else
          {
            loadUnits = 1;
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

          receptacles.Add(new ReceptacleLoadItem
          {
            ObjectId = receptacleId,
            Position = blockReference.Position,
            LoadUnits = loadUnits,
          });
        }
      }
      return receptacles;
    }

    private static List<ReceptacleCircuitGroup>
      BuildRoomReceptacleCircuitGroups(
        List<ReceptacleLoadItem> receptacles,
        int maximumLoadUnits)
    {
      int totalLoadUnits = 0;
      int largestReceptacleLoadUnits = 0;
      foreach (ReceptacleLoadItem receptacle in receptacles)
      {
        totalLoadUnits += receptacle.LoadUnits;
        largestReceptacleLoadUnits = Math.Max(
          largestReceptacleLoadUnits,
          receptacle.LoadUnits);
      }
      if (maximumLoadUnits < largestReceptacleLoadUnits)
      {
        throw new InvalidOperationException(
          $"The configured {maximumLoadUnits * ReceptacleLoadUnitKva:0.00} " +
          $"kVA circuit maximum is below a selected receptacle's " +
          $"{largestReceptacleLoadUnits * ReceptacleLoadUnitKva:0.00} kVA " +
          "load. Increase RC Max kVA in HRS.");
      }

      ReceptacleGroupingPlan bestPlan = null;
      for (int startIndex = 0;
           startIndex < receptacles.Count;
           startIndex++)
      {
        List<ReceptacleLoadItem> ordered =
          BuildNearestNeighborReceptacleOrder(receptacles, startIndex);
        ReceptacleGroupingPlan plan =
          PartitionOrderedReceptacles(
            ordered,
            maximumLoadUnits);
        if (IsBetterReceptacleGroupingPlan(plan, bestPlan))
        {
          bestPlan = plan;
        }
      }

      if (bestPlan == null || bestPlan.Groups.Count == 0)
      {
        throw new InvalidOperationException(
          $"The selected " +
          $"{totalLoadUnits * ReceptacleLoadUnitKva:0.00} kVA load cannot " +
          "be split into spatially contiguous circuits at or below the " +
          $"configured {maximumLoadUnits * ReceptacleLoadUnitKva:0.00} kVA " +
          "maximum while keeping quad receptacles together.");
      }

      bestPlan.Groups.Sort(CompareReceptacleCircuitGroupLocations);
      return bestPlan.Groups;
    }

    private static List<ReceptacleLoadItem>
      BuildNearestNeighborReceptacleOrder(
        List<ReceptacleLoadItem> receptacles,
        int startIndex)
    {
      List<ReceptacleLoadItem> ordered =
        new List<ReceptacleLoadItem>();
      bool[] used = new bool[receptacles.Count];
      int currentIndex = startIndex;

      while (ordered.Count < receptacles.Count)
      {
        ReceptacleLoadItem current = receptacles[currentIndex];
        ordered.Add(current);
        used[currentIndex] = true;

        int nextIndex = -1;
        double bestDistance = double.MaxValue;
        for (int candidateIndex = 0;
             candidateIndex < receptacles.Count;
             candidateIndex++)
        {
          if (used[candidateIndex])
          {
            continue;
          }

          double distance = ReceptacleDistanceSquared(
            current.Position,
            receptacles[candidateIndex].Position);
          if (distance < bestDistance - 1e-9 ||
              (Math.Abs(distance - bestDistance) <= 1e-9 &&
               (nextIndex < 0 ||
                CompareReceptacleLoadItems(
                  receptacles[candidateIndex],
                  receptacles[nextIndex]) < 0)))
          {
            bestDistance = distance;
            nextIndex = candidateIndex;
          }
        }

        if (nextIndex < 0)
        {
          break;
        }
        currentIndex = nextIndex;
      }
      return ordered;
    }

    private static ReceptacleGroupingPlan PartitionOrderedReceptacles(
      List<ReceptacleLoadItem> ordered,
      int maximumLoadUnits)
    {
      ReceptacleGroupingPlan[] bestFromIndex =
        new ReceptacleGroupingPlan[ordered.Count + 1];
      bestFromIndex[ordered.Count] = new ReceptacleGroupingPlan
      {
        Groups = new List<ReceptacleCircuitGroup>(),
        LoadBalanceScore = 0,
        SpatialCost = 0.0,
      };

      for (int start = ordered.Count - 1; start >= 0; start--)
      {
        int loadUnits = 0;
        for (int end = start; end < ordered.Count; end++)
        {
          loadUnits += ordered[end].LoadUnits;
          if (loadUnits > maximumLoadUnits)
          {
            break;
          }
          if (bestFromIndex[end + 1] == null)
          {
            continue;
          }

          List<ReceptacleLoadItem> groupItems =
            ordered.GetRange(start, end - start + 1);
          ReceptacleCircuitGroup group = new ReceptacleCircuitGroup
          {
            Items = groupItems,
            LoadUnits = loadUnits,
          };
          ReceptacleGroupingPlan suffix = bestFromIndex[end + 1];
          List<ReceptacleCircuitGroup> groups =
            new List<ReceptacleCircuitGroup> { group };
          groups.AddRange(suffix.Groups);

          ReceptacleGroupingPlan candidate = new ReceptacleGroupingPlan
          {
            Groups = groups,
            LoadBalanceScore = loadUnits * loadUnits +
              suffix.LoadBalanceScore,
            SpatialCost = CalculateReceptacleGroupSpatialCost(groupItems) +
              suffix.SpatialCost,
          };
          if (IsBetterReceptacleGroupingPlan(
            candidate,
            bestFromIndex[start]))
          {
            bestFromIndex[start] = candidate;
          }
        }
      }
      return bestFromIndex[0];
    }

    private static bool IsBetterReceptacleGroupingPlan(
      ReceptacleGroupingPlan candidate,
      ReceptacleGroupingPlan current)
    {
      if (candidate == null)
      {
        return false;
      }
      if (current == null)
      {
        return true;
      }
      if (candidate.Groups.Count != current.Groups.Count)
      {
        return candidate.Groups.Count < current.Groups.Count;
      }
      if (candidate.LoadBalanceScore != current.LoadBalanceScore)
      {
        return candidate.LoadBalanceScore < current.LoadBalanceScore;
      }
      return candidate.SpatialCost < current.SpatialCost - 1e-9;
    }

    private static double CalculateReceptacleGroupSpatialCost(
      List<ReceptacleLoadItem> receptacles)
    {
      Point3d center = CalculateReceptacleGroupCenter(receptacles);
      double cost = 0.0;
      foreach (ReceptacleLoadItem receptacle in receptacles)
      {
        cost += ReceptacleDistanceSquared(receptacle.Position, center);
      }
      return cost;
    }

    private static Point3d CalculateReceptacleGroupCenter(
      List<ReceptacleLoadItem> receptacles)
    {
      if (receptacles == null || receptacles.Count == 0)
      {
        return Point3d.Origin;
      }

      double x = 0.0;
      double y = 0.0;
      double z = 0.0;
      foreach (ReceptacleLoadItem receptacle in receptacles)
      {
        x += receptacle.Position.X;
        y += receptacle.Position.Y;
        z += receptacle.Position.Z;
      }
      return new Point3d(
        x / receptacles.Count,
        y / receptacles.Count,
        z / receptacles.Count);
    }

    private static int CompareReceptacleCircuitGroupLocations(
      ReceptacleCircuitGroup first,
      ReceptacleCircuitGroup second)
    {
      Point3d firstCenter = CalculateReceptacleGroupCenter(first.Items);
      Point3d secondCenter = CalculateReceptacleGroupCenter(second.Items);
      int verticalComparison = secondCenter.Y.CompareTo(firstCenter.Y);
      if (verticalComparison != 0)
      {
        return verticalComparison;
      }
      return firstCenter.X.CompareTo(secondCenter.X);
    }

    private static int CompareReceptacleLoadItems(
      ReceptacleLoadItem first,
      ReceptacleLoadItem second)
    {
      int xComparison = first.Position.X.CompareTo(second.Position.X);
      if (xComparison != 0)
      {
        return xComparison;
      }
      int yComparison = first.Position.Y.CompareTo(second.Position.Y);
      if (yComparison != 0)
      {
        return yComparison;
      }
      return first.ObjectId.Handle.Value.CompareTo(
        second.ObjectId.Handle.Value);
    }

    private static double ReceptacleDistanceSquared(
      Point3d first,
      Point3d second)
    {
      double x = first.X - second.X;
      double y = first.Y - second.Y;
      return x * x + y * y;
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
      double noteClearance =
        ReceptNoteClearanceInPaperInches * blockScale;
      double noteSearchStep = Math.Max(
        ReceptNoteSearchStepInPaperInches * blockScale,
        textHeight * 0.65);
      double noteMaximumGap = Math.Max(
        noteGap,
        ReceptNoteMaximumGapInPaperInches * blockScale);

      try
      {
        ObjectId noteLayerId = EnsureReceptNoteLayer(database);
        ObjectId textStyleId = EnsureHomerunTextStyle(database);
        int labeledCount = 0;
        int skippedCount = 0;
        int alternatePositionCount = 0;
        List<Extents3d> placedLabelExtents = new List<Extents3d>();

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
            if (PlaceReceptPanelLabelWithoutOverlap(
              editor,
              transaction,
              blockReference,
              panelLabel,
              noteDirection,
              noteGap,
              noteClearance,
              noteSearchStep,
              noteMaximumGap,
              placedLabelExtents))
            {
              alternatePositionCount++;
            }

            if (TryGetReceptLabelExtents(panelLabel, out Extents3d labelExtents))
            {
              placedLabelExtents.Add(labelExtents);
            }
            labeledCount++;
          }

          transaction.Commit();
        }

        editor.WriteMessage(
          $"\nAdded circuit label {panelCircuitLabel} to " +
          $"{labeledCount} receptacle(s)." +
          (alternatePositionCount > 0
            ? $" Placed {alternatePositionCount} label(s) in alternate " +
              "positions to avoid nearby objects."
            : string.Empty) +
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

    private static bool PlaceReceptPanelLabelWithoutOverlap(
      Editor editor,
      Transaction transaction,
      BlockReference blockReference,
      MText panelLabel,
      Vector3d preferredDirection,
      double gap,
      double clearance,
      double searchStep,
      double maximumGap,
      IList<Extents3d> additionalObstacleExtents)
    {
      double planarLength = Math.Sqrt(
        preferredDirection.X * preferredDirection.X +
        preferredDirection.Y * preferredDirection.Y);
      if (planarLength < 1e-9)
      {
        return false;
      }

      Vector3d forward = new Vector3d(
        preferredDirection.X / planarLength,
        preferredDirection.Y / planarLength,
        0.0);
      Vector3d[] directions = BuildReceptLabelCandidateDirections(forward);

      Point3d bestLocation = panelLabel.Location;
      double bestRotation = panelLabel.Rotation;
      AttachmentPoint bestAttachment = panelLabel.Attachment;
      double bestPenalty = double.MaxValue;
      int bestCandidateIndex = 0;
      int candidateIndex = 0;
      double safeSearchStep = Math.Max(searchStep, 1e-6);
      double safeMaximumGap = Math.Max(gap, maximumGap);
      int radialStepCount = Math.Max(
        0,
        (int)Math.Ceiling((safeMaximumGap - gap) / safeSearchStep));

      for (int stepIndex = 0;
           stepIndex <= radialStepCount;
           stepIndex++)
      {
        double candidateGap = Math.Min(
          safeMaximumGap,
          gap + stepIndex * safeSearchStep);
        for (int directionIndex = 0;
             directionIndex < directions.Length;
             directionIndex++)
        {
          Vector3d direction = directions[directionIndex];
          ResolveReceptNoteOrientation(
            Math.Atan2(direction.Y, direction.X),
            out double rotationDegrees,
            out AttachmentPoint attachment);

          panelLabel.Location = ResolveReceptNoteLocation(
            blockReference,
            blockReference.Position,
            direction,
            candidateGap);
          panelLabel.Rotation = rotationDegrees * Math.PI / 180.0;
          panelLabel.Attachment = attachment;

          if (!TryGetReceptLabelExtents(
            panelLabel,
            out Extents3d candidateExtents))
          {
            ApplyReceptLabelPlacement(
              panelLabel,
              bestLocation,
              bestRotation,
              bestAttachment);
            return false;
          }

          Extents3d clearanceExtents = ExpandReceptLabelExtents(
            candidateExtents,
            clearance);
          ReceptLabelObstacleScore obstacleScore =
            MeasureAdditionalReceptLabelObstacles(
            clearanceExtents,
            additionalObstacleExtents);
          obstacleScore.Add(MeasureDrawingReceptLabelObstacles(
            editor,
            transaction,
            clearanceExtents,
            blockReference.ObjectId,
            panelLabel.ObjectId));

          double distanceFraction = safeMaximumGap <= gap + 1e-9
            ? 0.0
            : (candidateGap - gap) / (safeMaximumGap - gap);
          double candidatePenalty = obstacleScore.BuildPenalty(
            clearanceExtents,
            distanceFraction,
            directionIndex);

          if (candidatePenalty < bestPenalty)
          {
            bestPenalty = candidatePenalty;
            bestLocation = panelLabel.Location;
            bestRotation = panelLabel.Rotation;
            bestAttachment = panelLabel.Attachment;
            bestCandidateIndex = candidateIndex;
          }

          if (obstacleScore.ObstacleCount == 0)
          {
            return candidateIndex != 0;
          }

          candidateIndex++;
        }
      }

      ApplyReceptLabelPlacement(
        panelLabel,
        bestLocation,
        bestRotation,
        bestAttachment);
      return bestCandidateIndex != 0;
    }

    private static Vector3d[] BuildReceptLabelCandidateDirections(
      Vector3d forward)
    {
      double eighthTurn = Math.PI / 4.0;
      double[] angleOffsets =
      {
        0.0,
        -eighthTurn,
        eighthTurn,
        -2.0 * eighthTurn,
        2.0 * eighthTurn,
        -3.0 * eighthTurn,
        3.0 * eighthTurn,
        Math.PI,
      };
      Vector3d[] directions = new Vector3d[angleOffsets.Length];
      for (int index = 0; index < angleOffsets.Length; index++)
      {
        directions[index] = forward.RotateBy(
          angleOffsets[index],
          Vector3d.ZAxis);
      }
      return directions;
    }

    private static void ApplyReceptLabelPlacement(
      MText panelLabel,
      Point3d location,
      double rotation,
      AttachmentPoint attachment)
    {
      panelLabel.Location = location;
      panelLabel.Rotation = rotation;
      panelLabel.Attachment = attachment;
    }

    private static bool TryGetReceptLabelExtents(
      MText panelLabel,
      out Extents3d extents)
    {
      try
      {
        extents = panelLabel.GeometricExtents;
        return true;
      }
      catch
      {
        extents = new Extents3d();
        return false;
      }
    }

    private static Extents3d ExpandReceptLabelExtents(
      Extents3d extents,
      double clearance)
    {
      double safeClearance = Math.Max(0.0, clearance);
      return new Extents3d(
        new Point3d(
          Math.Min(extents.MinPoint.X, extents.MaxPoint.X) - safeClearance,
          Math.Min(extents.MinPoint.Y, extents.MaxPoint.Y) - safeClearance,
          Math.Min(extents.MinPoint.Z, extents.MaxPoint.Z)),
        new Point3d(
          Math.Max(extents.MinPoint.X, extents.MaxPoint.X) + safeClearance,
          Math.Max(extents.MinPoint.Y, extents.MaxPoint.Y) + safeClearance,
          Math.Max(extents.MinPoint.Z, extents.MaxPoint.Z)));
    }

    private static ReceptLabelObstacleScore
      MeasureAdditionalReceptLabelObstacles(
      Extents3d candidateExtents,
      IList<Extents3d> additionalObstacleExtents)
    {
      ReceptLabelObstacleScore score = new ReceptLabelObstacleScore();
      if (additionalObstacleExtents == null)
      {
        return score;
      }

      foreach (Extents3d obstacleExtents in additionalObstacleExtents)
      {
        if (ReceptLabelExtentsIntersect(candidateExtents, obstacleExtents))
        {
          score.AddTextObstacle(
            CalculateReceptLabelIntersectionArea(
              candidateExtents,
              obstacleExtents));
        }
      }
      return score;
    }

    private static ReceptLabelObstacleScore
      MeasureDrawingReceptLabelObstacles(
      Editor editor,
      Transaction transaction,
      Extents3d candidateExtents,
      ObjectId receptacleId,
      ObjectId panelLabelId)
    {
      ReceptLabelObstacleScore score = new ReceptLabelObstacleScore();
      try
      {
        Matrix3d wcsToUcs = editor.CurrentUserCoordinateSystem.Inverse();
        double minX = Math.Min(
          candidateExtents.MinPoint.X,
          candidateExtents.MaxPoint.X);
        double maxX = Math.Max(
          candidateExtents.MinPoint.X,
          candidateExtents.MaxPoint.X);
        double minY = Math.Min(
          candidateExtents.MinPoint.Y,
          candidateExtents.MaxPoint.Y);
        double maxY = Math.Max(
          candidateExtents.MinPoint.Y,
          candidateExtents.MaxPoint.Y);
        double z = panelLabelId.IsNull
          ? candidateExtents.MinPoint.Z
          : 0.5 * (
            candidateExtents.MinPoint.Z + candidateExtents.MaxPoint.Z);

        Point3dCollection crossingPolygon = new Point3dCollection
        {
          new Point3d(minX, minY, z).TransformBy(wcsToUcs),
          new Point3d(maxX, minY, z).TransformBy(wcsToUcs),
          new Point3d(maxX, maxY, z).TransformBy(wcsToUcs),
          new Point3d(minX, maxY, z).TransformBy(wcsToUcs)
        };

        PromptSelectionResult selectionResult =
          editor.SelectCrossingPolygon(crossingPolygon);
        if (selectionResult.Status != PromptStatus.OK ||
            selectionResult.Value == null)
        {
          return score;
        }

        foreach (ObjectId selectedId in selectionResult.Value.GetObjectIds())
        {
          if (selectedId == receptacleId || selectedId == panelLabelId)
          {
            continue;
          }

          Entity entity = transaction.GetObject(
            selectedId,
            OpenMode.ForRead,
            false) as Entity;
          bool isText = entity is DBText ||
            entity is MText ||
            entity is MLeader;
          double overlapArea = 0.0;
          if (entity != null)
          {
            try
            {
              overlapArea = CalculateReceptLabelIntersectionArea(
                candidateExtents,
                entity.GeometricExtents);
            }
            catch
            {
            }
          }

          if (isText)
          {
            score.AddTextObstacle(overlapArea);
          }
          else
          {
            score.AddOtherObstacle(overlapArea);
          }
        }
        return score;
      }
      catch
      {
        // Preserve the original placement behavior if AutoCAD cannot perform
        // a crossing selection in the current view or coordinate system.
        return score;
      }
    }

    private static double CalculateReceptLabelIntersectionArea(
      Extents3d first,
      Extents3d second)
    {
      double overlapWidth = Math.Max(
        0.0,
        Math.Min(first.MaxPoint.X, second.MaxPoint.X) -
          Math.Max(first.MinPoint.X, second.MinPoint.X));
      double overlapHeight = Math.Max(
        0.0,
        Math.Min(first.MaxPoint.Y, second.MaxPoint.Y) -
          Math.Max(first.MinPoint.Y, second.MinPoint.Y));
      return overlapWidth * overlapHeight;
    }

    private static bool ReceptLabelExtentsIntersect(
      Extents3d first,
      Extents3d second)
    {
      double firstMinX = Math.Min(first.MinPoint.X, first.MaxPoint.X);
      double firstMaxX = Math.Max(first.MinPoint.X, first.MaxPoint.X);
      double firstMinY = Math.Min(first.MinPoint.Y, first.MaxPoint.Y);
      double firstMaxY = Math.Max(first.MinPoint.Y, first.MaxPoint.Y);
      double secondMinX = Math.Min(second.MinPoint.X, second.MaxPoint.X);
      double secondMaxX = Math.Max(second.MinPoint.X, second.MaxPoint.X);
      double secondMinY = Math.Min(second.MinPoint.Y, second.MaxPoint.Y);
      double secondMaxY = Math.Max(second.MinPoint.Y, second.MaxPoint.Y);

      return firstMinX <= secondMaxX &&
        firstMaxX >= secondMinX &&
        firstMinY <= secondMaxY &&
        firstMaxY >= secondMinY;
    }

    private sealed class ReceptLabelObstacleScore
    {
      internal int TextObstacleCount { get; private set; }
      internal int OtherObstacleCount { get; private set; }
      internal double TextOverlapArea { get; private set; }
      internal double OtherOverlapArea { get; private set; }

      internal int ObstacleCount =>
        TextObstacleCount + OtherObstacleCount;

      internal void AddTextObstacle(double overlapArea)
      {
        TextObstacleCount++;
        TextOverlapArea += Math.Max(0.0, overlapArea);
      }

      internal void AddOtherObstacle(double overlapArea)
      {
        OtherObstacleCount++;
        OtherOverlapArea += Math.Max(0.0, overlapArea);
      }

      internal void Add(ReceptLabelObstacleScore other)
      {
        if (other == null)
        {
          return;
        }
        TextObstacleCount += other.TextObstacleCount;
        OtherObstacleCount += other.OtherObstacleCount;
        TextOverlapArea += other.TextOverlapArea;
        OtherOverlapArea += other.OtherOverlapArea;
      }

      internal double BuildPenalty(
        Extents3d candidateExtents,
        double distanceFraction,
        int directionIndex)
      {
        double candidateArea = Math.Max(
          1e-9,
          Math.Abs(
            (candidateExtents.MaxPoint.X - candidateExtents.MinPoint.X) *
            (candidateExtents.MaxPoint.Y - candidateExtents.MinPoint.Y)));
        double overlapPenalty =
          (TextOverlapArea + OtherOverlapArea * 4.0) / candidateArea;
        return overlapPenalty * 100.0 +
          TextObstacleCount * 0.5 +
          OtherObstacleCount * 8.0 +
          Math.Max(0.0, distanceFraction) * 2.0 +
          Math.Max(0, directionIndex) * 0.02;
      }
    }

    private sealed class ReceptacleLoadItem
    {
      internal ObjectId ObjectId { get; set; }
      internal Point3d Position { get; set; }
      internal int LoadUnits { get; set; }
    }

    private sealed class ReceptacleCircuitGroup
    {
      internal List<ReceptacleLoadItem> Items { get; set; } =
        new List<ReceptacleLoadItem>();
      internal int LoadUnits { get; set; }

      internal ObjectId[] GetObjectIds()
      {
        ObjectId[] objectIds = new ObjectId[Items.Count];
        for (int index = 0; index < Items.Count; index++)
        {
          objectIds[index] = Items[index].ObjectId;
        }
        return objectIds;
      }
    }

    private sealed class ReceptacleGroupingPlan
    {
      internal List<ReceptacleCircuitGroup> Groups { get; set; } =
        new List<ReceptacleCircuitGroup>();
      internal int LoadBalanceScore { get; set; }
      internal double SpatialCost { get; set; }
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
