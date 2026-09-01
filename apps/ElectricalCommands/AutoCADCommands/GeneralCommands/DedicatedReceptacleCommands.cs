using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Globalization;
using System.Text;

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

    private static bool TryPromptDedicatedEquipment(
      Editor editor,
      out DedicatedEquipmentLoad equipment)
    {
      equipment = null;
      try
      {
        var window = new DedicatedEquipmentPickerWindow();
        bool? accepted =
          Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(
            window);
        if (accepted != true || window.SelectedEquipment == null)
        {
          editor.WriteMessage("\nDedicated circuiting canceled.");
          return false;
        }

        equipment = window.SelectedEquipment;
        return true;
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to open the dedicated-equipment picker: {ex.Message}");
        return false;
      }
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
            panelSchedule.SpareCount,
            new[]
            {
              new PanelScheduleCircuitRequest
              {
                ConnectedWatts = equipment.Kva * 1000.0,
                LoadDescription = equipment.Description,
                LoadTypeCode = "D",
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
          $"{equipment.Poles}P, {breakerAmps}A breaker." +
          (allocation.RemainingCounts != null
            ? $"\n{FormatPanelCircuitStatus(panelName, allocation.RemainingCounts)}"
            : string.Empty));

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

    internal static double CalculateDedicatedLoadAmps(
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

    internal static int SelectStandardDedicatedBreaker(double minimumAmps)
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
