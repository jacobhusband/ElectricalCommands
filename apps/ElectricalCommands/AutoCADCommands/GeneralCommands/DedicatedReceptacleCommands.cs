using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Globalization;
using System.Text;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string CompleteDedicatedCircuitCommand =
      "-RCCOMPLETEDEDICATED";

    private static readonly int[] StandardDedicatedBreakerSizes =
    {
      15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 100,
      110, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450,
      500, 600, 700, 800, 1000, 1200,
    };

    private static DedicatedEquipmentPickerWindow
      _activeDedicatedEquipmentPicker;
    private static PendingDedicatedCircuit _pendingDedicatedCircuit;

    private sealed class PendingDedicatedCircuit
    {
      internal Document Document { get; set; }

      internal ObjectId ReceptacleId { get; set; }

      internal double PaperInchesPerModelFoot { get; set; }

      internal DedicatedEquipmentLoad Equipment { get; set; }

      internal bool CompletionQueued { get; set; }
    }

    private static bool TryShowDedicatedEquipmentPicker(
      Document document,
      ObjectId receptacleId,
      double paperInchesPerModelFoot)
    {
      Editor editor = document.Editor;
      if (_activeDedicatedEquipmentPicker != null)
      {
        try
        {
          _activeDedicatedEquipmentPicker.Activate();
        }
        catch
        {
        }

        editor.WriteMessage(
          "\nA dedicated-circuit picker is already open. Complete or " +
          "cancel it before starting another dedicated circuit.");
        return false;
      }

      try
      {
        var window = new DedicatedEquipmentPickerWindow();
        _pendingDedicatedCircuit = new PendingDedicatedCircuit
        {
          Document = document,
          ReceptacleId = receptacleId,
          PaperInchesPerModelFoot = paperInchesPerModelFoot,
        };
        _activeDedicatedEquipmentPicker = window;
        window.CircuitAccepted += DedicatedEquipmentPicker_CircuitAccepted;
        window.Closed += DedicatedEquipmentPicker_Closed;
        AcApplication.ShowModelessWindow(window);

        editor.WriteMessage(
          "\nDedicated-circuit picker opened. You can continue moving " +
          "around AutoCAD; choose Use Circuit when the equipment details " +
          "are ready.");
        return true;
      }
      catch (System.Exception ex)
      {
        ClearDedicatedEquipmentPickerState();
        editor.WriteMessage(
          $"\nUnable to open the dedicated-equipment picker: {ex.Message}");
        return false;
      }
    }

    private static void DedicatedEquipmentPicker_CircuitAccepted(
      object sender,
      DedicatedEquipmentSelectedEventArgs e)
    {
      if (!(sender is DedicatedEquipmentPickerWindow window) ||
          window != _activeDedicatedEquipmentPicker ||
          _pendingDedicatedCircuit == null ||
          e?.Equipment == null)
      {
        return;
      }

      PendingDedicatedCircuit pending = _pendingDedicatedCircuit;
      pending.Equipment = e.Equipment;
      try
      {
        pending.Document.SendStringToExecute(
          CompleteDedicatedCircuitCommand + " ",
          true,
          false,
          false);
        pending.CompletionQueued = true;
      }
      catch (System.Exception ex)
      {
        pending.CompletionQueued = false;
        try
        {
          pending.Document.Editor.WriteMessage(
            $"\nUnable to queue the dedicated circuit: {ex.Message}");
        }
        catch
        {
        }
      }
    }

    private static void DedicatedEquipmentPicker_Closed(
      object sender,
      EventArgs e)
    {
      if (sender is DedicatedEquipmentPickerWindow window)
      {
        window.CircuitAccepted -= DedicatedEquipmentPicker_CircuitAccepted;
        window.Closed -= DedicatedEquipmentPicker_Closed;
      }

      _activeDedicatedEquipmentPicker = null;
      if (_pendingDedicatedCircuit == null ||
          _pendingDedicatedCircuit.CompletionQueued)
      {
        return;
      }

      try
      {
        _pendingDedicatedCircuit.Document.Editor.WriteMessage(
          "\nDedicated circuiting canceled.");
      }
      catch
      {
      }
      _pendingDedicatedCircuit = null;
    }

    private static void ClearDedicatedEquipmentPickerState()
    {
      if (_activeDedicatedEquipmentPicker != null)
      {
        _activeDedicatedEquipmentPicker.CircuitAccepted -=
          DedicatedEquipmentPicker_CircuitAccepted;
        _activeDedicatedEquipmentPicker.Closed -=
          DedicatedEquipmentPicker_Closed;
      }
      _activeDedicatedEquipmentPicker = null;
      _pendingDedicatedCircuit = null;
    }

    [CommandMethod(
      CompleteDedicatedCircuitCommand,
      CommandFlags.Modal | CommandFlags.NoHistory)]
    public static void CompleteDedicatedReceptacleCircuit()
    {
      PendingDedicatedCircuit pending = _pendingDedicatedCircuit;
      _pendingDedicatedCircuit = null;
      if (pending?.Equipment == null)
      {
        return;
      }

      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        return;
      }

      Editor editor = document.Editor;
      if (pending.Document != document)
      {
        editor.WriteMessage(
          "\nDedicated circuiting canceled because the source drawing " +
          "is no longer active.");
        return;
      }

      Database database = document.Database;
      if (!TryValidatePendingDedicatedReceptacle(
        database,
        pending.ReceptacleId))
      {
        editor.WriteMessage(
          "\nDedicated circuiting canceled because the selected " +
          "receptacle is no longer available.");
        return;
      }

      if (!ElectricalDrawingSettingsStore.TryReadPanelName(
        database,
        out string panelName))
      {
        editor.WriteMessage(
          "\nDedicated circuiting requires a panel name. " +
          "Run SETPANELNAME (SPN) first.");
        return;
      }

      if (!ElectricalDrawingSettingsStore.TryReadPanelSchedule(
          database,
          out var panelSchedule) ||
        !System.IO.File.Exists(panelSchedule.WorkbookPath))
      {
        editor.WriteMessage(
          "\nDedicated circuiting requires a linked panel schedule. " +
          "Run SETPANELSCHEDULE (SPS) first.");
        return;
      }

      if (!TryVerifyPanelScheduleWorkbookClosed(
        panelSchedule.WorkbookPath,
        out string workbookAvailabilityError))
      {
        editor.WriteMessage(
          $"\nDedicated circuiting canceled: {workbookAvailabilityError}");
        return;
      }

      AutomaticallyCircuitDedicatedReceptacle(
        database,
        editor,
        pending.ReceptacleId,
        pending.PaperInchesPerModelFoot,
        panelName,
        pending.Equipment);
    }

    private static bool TryValidatePendingDedicatedReceptacle(
      Database database,
      ObjectId receptacleId)
    {
      if (receptacleId.IsNull || receptacleId.IsErased)
      {
        return false;
      }

      try
      {
        using (Transaction transaction =
          database.TransactionManager.StartOpenCloseTransaction())
        {
          BlockReference blockReference = transaction.GetObject(
            receptacleId,
            OpenMode.ForRead,
            false) as BlockReference;
          return blockReference != null &&
            IsSupportedReceptacleBlock(transaction, blockReference);
        }
      }
      catch
      {
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
