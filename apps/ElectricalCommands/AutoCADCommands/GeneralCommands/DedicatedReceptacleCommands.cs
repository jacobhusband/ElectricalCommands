using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private static readonly int[] StandardDedicatedBreakerSizes =
    {
      15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 100,
      110, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450,
      500, 600, 700, 800, 1000, 1200,
    };

    private static bool TryPromptDedicatedReceptacle(
      Database database,
      Editor editor,
      out ObjectId receptacleId)
    {
      receptacleId = ObjectId.Null;
      PromptEntityOptions options = new PromptEntityOptions(
        $"\nSelect one {ReceptBlockName} or " +
        $"{AlternateReceptBlockName} block for the dedicated circuit: ");
      options.SetRejectMessage("\nSelect a receptacle block.");
      options.AddAllowedClass(typeof(BlockReference), false);

      PromptEntityResult result = editor.GetEntity(options);
      if (result.Status != PromptStatus.OK)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }

      try
      {
        using (Transaction transaction =
          database.TransactionManager.StartOpenCloseTransaction())
        {
          BlockReference blockReference = transaction.GetObject(
            result.ObjectId,
            OpenMode.ForRead,
            false) as BlockReference;
          if (blockReference == null ||
              blockReference.OwnerId != database.CurrentSpaceId ||
              !IsSupportedReceptacleBlock(transaction, blockReference))
          {
            editor.WriteMessage(
              $"\nThe selected object is not a supported receptacle block " +
              "in the current drawing space.");
            return false;
          }
        }
        receptacleId = result.ObjectId;
        return true;
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to read the selected receptacle: {ex.Message}");
        return false;
      }
    }

    private static bool TryPromptDedicatedEquipment(
      Editor editor,
      out DedicatedEquipmentLoad equipment)
    {
      equipment = null;
      PromptStringOptions descriptionOptions = new PromptStringOptions(
        "\nEnter dedicated equipment description " +
        "(for example FRIDGE, MICROWAVE, COUNTER, or WASHER): ")
      {
        AllowSpaces = true,
      };
      PromptResult descriptionResult = editor.GetString(descriptionOptions);
      if (descriptionResult.Status != PromptStatus.OK)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }

      string description = Regex.Replace(
        descriptionResult.StringResult ?? string.Empty,
        @"\s+",
        " ").Trim().ToUpperInvariant();
      if (description.Length == 0)
      {
        editor.WriteMessage("\nEquipment description cannot be blank.");
        return false;
      }

      DedicatedEquipmentLoad preset = null;
      string catalogPath = string.Empty;
      try
      {
        DedicatedEquipmentCatalog.TryFind(
          description,
          out preset,
          out catalogPath);
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nEquipment catalog warning: {ex.Message} " +
          "Enter custom equipment values instead.");
      }

      if (preset != null)
      {
        editor.WriteMessage(
          $"\nMatched {preset.Description} preset: {preset.Kva:0.###} kVA, " +
          $"{preset.Voltage}V, {preset.Poles}P, " +
          $"{ResolveDedicatedBreakerAmps(preset)}A breaker.");
        PromptKeywordOptions sourceOptions = new PromptKeywordOptions(
          "\nUse catalog values or enter equipment-schedule values " +
          "[Preset/Custom] <Preset>: ",
          "Preset Custom")
        {
          AllowNone = true,
        };
        PromptResult sourceResult = editor.GetKeywords(sourceOptions);
        if (sourceResult.Status != PromptStatus.OK &&
            sourceResult.Status != PromptStatus.None)
        {
          editor.WriteMessage("\nDedicated circuiting canceled.");
          return false;
        }
        bool usePreset = sourceResult.Status == PromptStatus.None ||
          string.Equals(
            sourceResult.StringResult,
            "Preset",
            StringComparison.OrdinalIgnoreCase);
        if (usePreset)
        {
          equipment = CopyDedicatedEquipment(preset, description);
          editor.WriteMessage(
            $"\nUsing dedicated-equipment catalog at \"{catalogPath}\".");
          return true;
        }
      }

      return TryPromptCustomDedicatedEquipment(
        editor,
        description,
        out equipment);
    }

    private static bool TryPromptCustomDedicatedEquipment(
      Editor editor,
      string description,
      out DedicatedEquipmentLoad equipment)
    {
      equipment = null;
      PromptDoubleOptions kvaOptions = new PromptDoubleOptions(
        "\nEnter connected equipment load in kVA: ")
      {
        AllowNegative = false,
        AllowZero = false,
      };
      PromptDoubleResult kvaResult = editor.GetDouble(kvaOptions);
      if (kvaResult.Status != PromptStatus.OK)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }

      PromptIntegerOptions voltageOptions = new PromptIntegerOptions(
        "\nEnter equipment voltage <120>: ")
      {
        AllowNone = true,
        AllowNegative = false,
        AllowZero = false,
        DefaultValue = 120,
        LowerLimit = 1,
        UpperLimit = 1000,
        UseDefaultValue = true,
      };
      PromptIntegerResult voltageResult = editor.GetInteger(voltageOptions);
      if (voltageResult.Status != PromptStatus.OK &&
          voltageResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }
      int voltage = voltageResult.Status == PromptStatus.OK
        ? voltageResult.Value
        : 120;

      int defaultPoles = voltage <= 120 ? 1 : 2;
      PromptIntegerOptions poleOptions = new PromptIntegerOptions(
        $"\nEnter circuit poles <{defaultPoles}>: ")
      {
        AllowNone = true,
        AllowNegative = false,
        AllowZero = false,
        DefaultValue = defaultPoles,
        LowerLimit = 1,
        UpperLimit = 3,
        UseDefaultValue = true,
      };
      PromptIntegerResult poleResult = editor.GetInteger(poleOptions);
      if (poleResult.Status != PromptStatus.OK &&
          poleResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }
      int poles = poleResult.Status == PromptStatus.OK
        ? poleResult.Value
        : defaultPoles;

      PromptDoubleOptions mcaOptions = new PromptDoubleOptions(
        "\nEnter MCA in amps <not specified>: ")
      {
        AllowNone = true,
        AllowNegative = false,
        AllowZero = false,
      };
      PromptDoubleResult mcaResult = editor.GetDouble(mcaOptions);
      if (mcaResult.Status != PromptStatus.OK &&
          mcaResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }
      double? mca = mcaResult.Status == PromptStatus.OK
        ? (double?)mcaResult.Value
        : null;

      double calculatedAmps = CalculateDedicatedLoadAmps(
        kvaResult.Value,
        voltage,
        poles);
      int defaultBreaker = SelectStandardDedicatedBreaker(
        Math.Max(calculatedAmps, mca ?? 0.0));
      PromptIntegerOptions mocpOptions = new PromptIntegerOptions(
        $"\nEnter MOCP / circuit-breaker amps <{defaultBreaker}>: ")
      {
        AllowNone = true,
        AllowNegative = false,
        AllowZero = false,
        DefaultValue = defaultBreaker,
        LowerLimit = 1,
        UpperLimit = 1200,
        UseDefaultValue = true,
      };
      PromptIntegerResult mocpResult = editor.GetInteger(mocpOptions);
      if (mocpResult.Status != PromptStatus.OK &&
          mocpResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }
      int mocp = mocpResult.Status == PromptStatus.OK
        ? mocpResult.Value
        : defaultBreaker;
      double minimumBreakerAmps = Math.Max(
        calculatedAmps,
        mca ?? 0.0);
      if (mocp + 1e-9 < minimumBreakerAmps)
      {
        editor.WriteMessage(
          $"\nThe {mocp}A MOCP is below the required " +
          $"{minimumBreakerAmps:0.##}A load/MCA. Dedicated circuiting " +
          "canceled so the equipment values can be corrected.");
        return false;
      }

      equipment = new DedicatedEquipmentLoad
      {
        Description = description,
        Kva = kvaResult.Value,
        Voltage = voltage,
        Poles = poles,
        McaAmps = mca,
        MocpAmps = mocp,
        LoadTypeCode = "D",
      };

      PromptKeywordOptions saveOptions = new PromptKeywordOptions(
        "\nSave these values as a reusable equipment preset " +
        "[Yes/No] <No>: ",
        "Yes No")
      {
        AllowNone = true,
      };
      PromptResult saveResult = editor.GetKeywords(saveOptions);
      if (saveResult.Status != PromptStatus.OK &&
          saveResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nDedicated circuiting canceled.");
        return false;
      }
      if (saveResult.Status == PromptStatus.OK &&
          string.Equals(
            saveResult.StringResult,
            "Yes",
            StringComparison.OrdinalIgnoreCase))
      {
        try
        {
          string catalogPath = DedicatedEquipmentCatalog.SaveOrUpdate(
            equipment);
          editor.WriteMessage(
            $"\nSaved dedicated-equipment preset to \"{catalogPath}\".");
        }
        catch (System.Exception ex)
        {
          editor.WriteMessage(
            $"\nCould not save the equipment preset: {ex.Message}");
        }
      }
      return true;
    }

    private static void AutomaticallyCircuitDedicatedReceptacle(
      Database database,
      Editor editor,
      ObjectId receptacleId,
      double paperInchesPerModelFoot,
      string panelName,
      DedicatedEquipmentLoad equipment)
    {
      if (!ElectricalDrawingSettingsStore.TryReadPanelSchedule(
        database,
        out var panelSchedule))
      {
        editor.WriteMessage(
          "\nDedicated circuiting requires a linked panel schedule. " +
          "Run SETPANELSCHEDULE (SPS) first.");
        return;
      }

      int breakerAmps = ResolveDedicatedBreakerAmps(equipment);
      try
      {
        PanelScheduleAllocationResult allocation =
          PanelScheduleWorkbookAllocator.AllocateReceptacleCircuits(
            panelSchedule.WorkbookPath,
            panelName,
            panelSchedule.CircuitCapacity,
            new[]
            {
              new PanelScheduleCircuitRequest
              {
                ConnectedWatts = equipment.Kva * 1000.0,
                LoadDescription = equipment.Description,
                LoadTypeCode = "D",
                Notes = BuildDedicatedPanelNotes(equipment),
                Poles = equipment.Poles,
                BreakerAmps = breakerAmps,
              },
            })[0];

        AddCircuitLabelsToReceptacles(
          database,
          editor,
          new[] { receptacleId },
          paperInchesPerModelFoot,
          panelName,
          allocation.CircuitLabel);

        editor.WriteMessage(
          $"\nDedicated circuit added to worksheet " +
          $"\"{allocation.WorksheetName}\", circuit " +
          $"{allocation.CircuitLabel}: {equipment.Description}, " +
          $"{equipment.Kva:0.###} kVA, {equipment.Voltage}V, " +
          $"{equipment.Poles}P, {breakerAmps}A breaker.");

        OfferDedicatedKeyedNoteSource(
          database,
          editor,
          equipment,
          breakerAmps);
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to add the dedicated circuit: {ex.Message}");
      }
    }

    private static void OfferDedicatedKeyedNoteSource(
      Database database,
      Editor editor,
      DedicatedEquipmentLoad equipment,
      int breakerAmps)
    {
      PromptKeywordOptions noteOptions = new PromptKeywordOptions(
        "\nCreate standard equipment keyed-note source text " +
        "[Yes/No] <No>: ",
        "Yes No")
      {
        AllowNone = true,
      };
      PromptResult noteResult = editor.GetKeywords(noteOptions);
      if (noteResult.Status != PromptStatus.OK ||
          !string.Equals(
            noteResult.StringResult,
            "Yes",
            StringComparison.OrdinalIgnoreCase))
      {
        return;
      }

      PromptPointOptions pointOptions = new PromptPointOptions(
        "\nSpecify insertion point for equipment keyed-note source text: ")
      {
        AllowNone = false,
      };
      PromptPointResult pointResult = editor.GetPoint(pointOptions);
      if (pointResult.Status != PromptStatus.OK)
      {
        editor.WriteMessage("\nEquipment keyed-note creation canceled.");
        return;
      }

      AddNoteToDrawing(
        database,
        editor,
        pointResult.Value,
        BuildDedicatedKeyedNote(equipment, breakerAmps));
      editor.WriteMessage(
        "\nEquipment note created. Select it with KNT to add it to the " +
        "keyed-note table.");
    }

    private static string BuildDedicatedPanelNotes(
      DedicatedEquipmentLoad equipment)
    {
      var parts = new StringBuilder();
      parts.Append(equipment.Voltage.ToString(CultureInfo.InvariantCulture));
      parts.Append("V");
      if (equipment.McaAmps.HasValue)
      {
        parts.Append("; MCA ");
        parts.Append(equipment.McaAmps.Value.ToString(
          "0.##",
          CultureInfo.InvariantCulture));
        parts.Append("A");
      }
      if (equipment.MocpAmps.HasValue)
      {
        parts.Append("; MOCP ");
        parts.Append(equipment.MocpAmps.Value.ToString(
          CultureInfo.InvariantCulture));
        parts.Append("A");
      }
      return parts.ToString();
    }

    private static string BuildDedicatedKeyedNote(
      DedicatedEquipmentLoad equipment,
      int breakerAmps)
    {
      var note = new StringBuilder();
      note.Append("PROVIDE DEDICATED ");
      note.Append(equipment.Voltage.ToString(CultureInfo.InvariantCulture));
      note.Append("V, ");
      note.Append(equipment.Poles.ToString(CultureInfo.InvariantCulture));
      note.Append("-POLE, ");
      note.Append(breakerAmps.ToString(CultureInfo.InvariantCulture));
      note.Append("A CIRCUIT FOR ");
      note.Append(equipment.Description);
      note.Append(" (");
      note.Append(equipment.Kva.ToString("0.###", CultureInfo.InvariantCulture));
      note.Append(" KVA");
      if (equipment.McaAmps.HasValue)
      {
        note.Append(", MCA ");
        note.Append(equipment.McaAmps.Value.ToString(
          "0.##",
          CultureInfo.InvariantCulture));
        note.Append("A");
      }
      if (equipment.MocpAmps.HasValue)
      {
        note.Append(", MOCP ");
        note.Append(equipment.MocpAmps.Value.ToString(
          CultureInfo.InvariantCulture));
        note.Append("A");
      }
      note.Append("). CONNECT PER MANUFACTURER'S REQUIREMENTS. VERIFY FINAL ");
      note.Append("LOCATION AND ELECTRICAL REQUIREMENTS WITH THE EQUIPMENT ");
      note.Append("SUPPLIER PRIOR TO ROUGH-IN.");
      return note.ToString();
    }

    private static DedicatedEquipmentLoad CopyDedicatedEquipment(
      DedicatedEquipmentLoad source,
      string description)
    {
      return new DedicatedEquipmentLoad
      {
        Description = description,
        Kva = source.Kva,
        Voltage = source.Voltage,
        Poles = source.Poles,
        McaAmps = source.McaAmps,
        MocpAmps = source.MocpAmps,
        LoadTypeCode = source.LoadTypeCode,
      };
    }

    private static int ResolveDedicatedBreakerAmps(
      DedicatedEquipmentLoad equipment)
    {
      if (equipment.MocpAmps.HasValue && equipment.MocpAmps.Value > 0)
      {
        return equipment.MocpAmps.Value;
      }
      double loadAmps = CalculateDedicatedLoadAmps(
        equipment.Kva,
        equipment.Voltage,
        equipment.Poles);
      return SelectStandardDedicatedBreaker(
        Math.Max(loadAmps, equipment.McaAmps ?? 0.0));
    }

    private static double CalculateDedicatedLoadAmps(
      double kva,
      int voltage,
      int poles)
    {
      if (kva <= 0.0 || voltage <= 0)
      {
        return 0.0;
      }
      double denominator = voltage;
      if (poles == 3)
      {
        denominator *= Math.Sqrt(3.0);
      }
      return kva * 1000.0 / denominator;
    }

    private static int SelectStandardDedicatedBreaker(double minimumAmps)
    {
      foreach (int breakerSize in StandardDedicatedBreakerSizes)
      {
        if (breakerSize + 1e-9 >= minimumAmps)
        {
          return breakerSize;
        }
      }
      throw new InvalidOperationException(
        $"The calculated {minimumAmps:0.##}A load exceeds the supported " +
        "1200A breaker range.");
    }
  }
}
