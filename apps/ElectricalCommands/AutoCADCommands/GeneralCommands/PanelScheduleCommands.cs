using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Excel = Microsoft.Office.Interop.Excel;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const int DefaultPanelCircuitCapacity = 42;
    internal const int ReservedPanelCircuitCount = 6;

    [CommandMethod("SETPANELSCHEDULE", CommandFlags.Modal)]
    [CommandMethod("SPS", CommandFlags.Modal)]
    public static void SetPanelScheduleCommand()
    {
      Document document = Application.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        return;
      }

      Editor editor = document.Editor;
      if (!ElectricalDrawingSettingsStore.TryReadPanelName(
        document.Database,
        out string panelName))
      {
        editor.WriteMessage(
          "\nSet the panel name with SETPANELNAME (SPN) before " +
          "linking a panel schedule.");
        return;
      }

      ElectricalDrawingSettingsStore.TryReadPanelSchedule(
        document.Database,
        out var existingSetting);

      var dialog = new Microsoft.Win32.OpenFileDialog
      {
        CheckFileExists = true,
        DefaultExt = ".xls",
        Filter = "Excel panel schedules (*.xls;*.xlsx)|*.xls;*.xlsx",
        Multiselect = false,
        Title = $"Select the panel schedule workbook for {panelName}",
      };
      if (existingSetting != null &&
          File.Exists(existingSetting.WorkbookPath))
      {
        dialog.FileName = existingSetting.WorkbookPath;
        dialog.InitialDirectory = Path.GetDirectoryName(
          existingSetting.WorkbookPath);
      }

      if (dialog.ShowDialog() != true)
      {
        editor.WriteMessage("\nPanel schedule selection canceled.");
        return;
      }

      int defaultCapacity = existingSetting?.CircuitCapacity ??
        DefaultPanelCircuitCapacity;
      PromptIntegerOptions capacityOptions = new PromptIntegerOptions(
        $"\nEnter the number of circuits in panel {panelName} " +
        $"<{defaultCapacity}>: ")
      {
        AllowNone = true,
        AllowNegative = false,
        AllowZero = false,
        DefaultValue = defaultCapacity,
        LowerLimit = ReservedPanelCircuitCount,
        UpperLimit = 84,
        UseDefaultValue = true,
      };
      PromptIntegerResult capacityResult = editor.GetInteger(capacityOptions);
      if (capacityResult.Status != PromptStatus.OK &&
          capacityResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nPanel schedule selection canceled.");
        return;
      }

      int circuitCapacity = capacityResult.Status == PromptStatus.None
        ? defaultCapacity
        : capacityResult.Value;
      if (circuitCapacity % 2 != 0)
      {
        editor.WriteMessage("\nPanel circuit count must be an even number.");
        return;
      }

      try
      {
        string worksheetName = PanelScheduleWorkbookAllocator.PreparePanel(
          dialog.FileName,
          panelName,
          circuitCapacity);
        ElectricalDrawingSettingsStore.WritePanelSchedule(
          document.Database,
          Path.GetFullPath(dialog.FileName),
          circuitCapacity);
        HomerunSettingsPalette.Refresh();

        editor.WriteMessage(
          $"\nPanel schedule linked to worksheet \"{worksheetName}\". " +
          $"Circuits 1-{circuitCapacity} are active; " +
          $"{BuildReservedCircuitSummary(circuitCapacity)} are reserved as spares.");
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to link the panel schedule: {ex.Message}");
      }
    }

    private static string BuildReservedCircuitSummary(int circuitCapacity)
    {
      List<string> numbers = new List<string>();
      for (int circuit = circuitCapacity - ReservedPanelCircuitCount + 1;
           circuit <= circuitCapacity;
           circuit++)
      {
        numbers.Add(circuit.ToString(CultureInfo.InvariantCulture));
      }
      return "circuits " + string.Join(", ", numbers);
    }
  }

  internal sealed class PanelScheduleAllocationResult
  {
    internal int CircuitNumber { get; set; }
    internal string WorksheetName { get; set; } = string.Empty;
    internal double ConnectedWatts { get; set; }
  }

  internal static class PanelScheduleWorkbookAllocator
  {
    private const int AciesStartRow = 7;
    private const int AciesMaxRow = 48;
    private const int GeneratedStartRow = 8;
    private const int GeneratedMaxRow = 28;

    internal static string PreparePanel(
      string workbookPath,
      string panelName,
      int circuitCapacity)
    {
      return ExecuteWithPanelWorksheet(
        workbookPath,
        panelName,
        worksheet =>
        {
          List<PanelCircuitSlot> slots = BuildCircuitSlots(worksheet);
          ApplyCapacityAndSpareRules(worksheet, slots, circuitCapacity);
          return worksheet.Name;
        });
    }

    internal static PanelScheduleAllocationResult AllocateReceptacleCircuit(
      string workbookPath,
      string panelName,
      int circuitCapacity,
      double connectedWatts)
    {
      if (connectedWatts <= 0.0 ||
          double.IsNaN(connectedWatts) ||
          double.IsInfinity(connectedWatts))
      {
        throw new InvalidOperationException(
          "The selected receptacle load must be greater than zero.");
      }

      return ExecuteWithPanelWorksheet(
        workbookPath,
        panelName,
        worksheet =>
        {
          List<PanelCircuitSlot> slots = BuildCircuitSlots(worksheet);
          ApplyCapacityAndSpareRules(worksheet, slots, circuitCapacity);

          PanelCircuitSlot available = FindAvailableSlot(
            worksheet,
            slots,
            circuitCapacity);
          if (available == null)
          {
            throw new InvalidOperationException(
              $"Panel {panelName} has no usable SPARE, SPACE, or empty " +
              "circuit before its six reserved spares.");
          }

          WriteCell(worksheet, available.LoadDescriptionAddress, "RECEPTACLES");
          WriteCell(worksheet, available.LoadTypeAddress, "G");
          WriteCell(worksheet, available.PolesAddress, 1);
          WriteCell(worksheet, available.BreakerAmpsAddress, 20);
          WriteCell(
            worksheet,
            available.ConnectedKvaAddress,
            Math.Round(connectedWatts / 1000.0, 3));

          return new PanelScheduleAllocationResult
          {
            CircuitNumber = available.CircuitNumber,
            WorksheetName = worksheet.Name,
            ConnectedWatts = connectedWatts,
          };
        });
    }

    private static T ExecuteWithPanelWorksheet<T>(
      string workbookPath,
      string panelName,
      Func<Excel.Worksheet, T> action)
    {
      string fullPath = Path.GetFullPath(workbookPath ?? string.Empty);
      if (!File.Exists(fullPath))
      {
        throw new FileNotFoundException(
          "The linked panel schedule workbook could not be found.",
          fullPath);
      }

      string extension = Path.GetExtension(fullPath);
      if (!string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) &&
          !string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
      {
        throw new InvalidOperationException(
          "Panel schedule must be an .xls or .xlsx workbook.");
      }

      Excel.Application excel = null;
      Excel.Workbooks workbooks = null;
      Excel.Workbook workbook = null;
      Excel.Sheets worksheets = null;
      Excel.Worksheet worksheet = null;

      try
      {
        excel = new Excel.Application
        {
          Visible = false,
          DisplayAlerts = false,
          AskToUpdateLinks = false,
        };
        try
        {
          excel.AutomationSecurity =
            Microsoft.Office.Core.MsoAutomationSecurity
              .msoAutomationSecurityForceDisable;
        }
        catch
        {
          // Older Excel versions may not expose AutomationSecurity.
        }
        workbooks = excel.Workbooks;
        workbook = workbooks.Open(
          fullPath,
          UpdateLinks: 0,
          ReadOnly: false,
          IgnoreReadOnlyRecommended: true,
          AddToMru: false);
        if (workbook.ReadOnly)
        {
          throw new InvalidOperationException(
            "The panel schedule is read-only or open elsewhere. Close Excel " +
            "and try again.");
        }

        worksheets = workbook.Worksheets;
        worksheet = FindPanelWorksheet(worksheets, panelName);
        T result = action(worksheet);

        try
        {
          excel.Calculate();
        }
        catch
        {
          // Workbook formulas will calculate the next time Excel opens it.
        }
        workbook.Save();
        return result;
      }
      catch (COMException ex)
      {
        throw new InvalidOperationException(
          "Microsoft Excel could not update the panel schedule. " + ex.Message,
          ex);
      }
      finally
      {
        if (workbook != null)
        {
          try
          {
            workbook.Close(false);
          }
          catch
          {
          }
        }
        if (excel != null)
        {
          try
          {
            excel.Quit();
          }
          catch
          {
          }
        }

        ReleaseComObject(worksheet);
        ReleaseComObject(worksheets);
        ReleaseComObject(workbook);
        ReleaseComObject(workbooks);
        ReleaseComObject(excel);
      }
    }

    private static Excel.Worksheet FindPanelWorksheet(
      Excel.Sheets worksheets,
      string panelName)
    {
      string requestedName = NormalizePanelName(panelName);
      List<string> scheduleNames = new List<string>();

      for (int index = 1; index <= worksheets.Count; index++)
      {
        Excel.Worksheet candidate = null;
        bool keepCandidate = false;
        try
        {
          candidate = worksheets.Item[index] as Excel.Worksheet;
          if (candidate == null || !IsPanelSchedule(candidate))
          {
            continue;
          }

          scheduleNames.Add(candidate.Name);
          if (MatchesPanelName(candidate, requestedName))
          {
            keepCandidate = true;
            return candidate;
          }
        }
        finally
        {
          if (!keepCandidate)
          {
            ReleaseComObject(candidate);
          }
        }
      }

      throw new InvalidOperationException(
        $"Could not find panel {panelName} in the workbook. Available panel " +
        $"worksheets: {(scheduleNames.Count == 0 ? "none" : string.Join(", ", scheduleNames))}.");
    }

    private static bool IsPanelSchedule(Excel.Worksheet worksheet)
    {
      string attachedHeader = ReadText(worksheet, "A5");
      string generatedHeader = ReadText(worksheet, "A6");
      return string.Equals(
          attachedHeader,
          "CKT#",
          StringComparison.OrdinalIgnoreCase) ||
        string.Equals(
          generatedHeader,
          "CKT#",
          StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesPanelName(
      Excel.Worksheet worksheet,
      string requestedName)
    {
      string[] candidates =
      {
        worksheet.Name,
        ReadText(worksheet, "A2"),
        ReadText(worksheet, "A3"),
      };
      foreach (string candidate in candidates)
      {
        if (string.Equals(
          NormalizePanelName(candidate),
          requestedName,
          StringComparison.OrdinalIgnoreCase))
        {
          return true;
        }
      }
      return false;
    }

    private static string NormalizePanelName(string value)
    {
      string text = (value ?? string.Empty).Trim();
      Match quoted = Regex.Match(
        text,
        @"\b(?:PANEL|PNL)\s*[\""']([^\""']+)[\""']",
        RegexOptions.IgnoreCase);
      if (quoted.Success)
      {
        return quoted.Groups[1].Value.Trim().ToUpperInvariant();
      }

      text = Regex.Replace(text, @"^\s*\([^)]*\)\s*", string.Empty);
      text = Regex.Replace(
        text,
        @"^\s*(?:PANEL|PNL)\s*",
        string.Empty,
        RegexOptions.IgnoreCase);
      return text.Trim(' ', '\"', '\'').ToUpperInvariant();
    }

    private static List<PanelCircuitSlot> BuildCircuitSlots(
      Excel.Worksheet worksheet)
    {
      if (string.Equals(
        ReadText(worksheet, "A5"),
        "CKT#",
        StringComparison.OrdinalIgnoreCase))
      {
        return BuildAttachedAciesSlots(worksheet);
      }

      List<PanelCircuitSlot> slots = new List<PanelCircuitSlot>();
      for (int row = GeneratedStartRow; row <= GeneratedMaxRow; row++)
      {
        int oddCircuit = (row - GeneratedStartRow) * 2 + 1;
        slots.Add(new PanelCircuitSlot
        {
          CircuitNumber = oddCircuit,
          Row = row,
          LoadTypeAddress = $"C{row}",
          PolesAddress = $"D{row}",
          BreakerAmpsAddress = $"E{row}",
          LoadDescriptionAddress = $"F{row}",
          ConnectedKvaAddress = $"I{row}",
        });
        slots.Add(new PanelCircuitSlot
        {
          CircuitNumber = oddCircuit + 1,
          Row = row,
          ConnectedKvaAddress = $"K{row}",
          LoadDescriptionAddress = $"L{row}",
          PolesAddress = $"O{row}",
          BreakerAmpsAddress = $"P{row}",
          LoadTypeAddress = $"Q{row}",
        });
      }
      return slots;
    }

    private static List<PanelCircuitSlot> BuildAttachedAciesSlots(
      Excel.Worksheet worksheet)
    {
      List<PanelCircuitSlot> slots = new List<PanelCircuitSlot>();
      for (int row = AciesStartRow; row <= AciesMaxRow; row++)
      {
        int leftCircuit = ReadInteger(worksheet, $"A{row}");
        if (leftCircuit > 0)
        {
          slots.Add(new PanelCircuitSlot
          {
            CircuitNumber = leftCircuit,
            Row = row,
            LoadTypeAddress = $"C{row}",
            PolesAddress = $"D{row}",
            BreakerAmpsAddress = $"E{row}",
            LoadDescriptionAddress = $"F{row}",
            ConnectedKvaAddress = $"K{row}",
          });
        }

        int rightCircuit = ReadInteger(worksheet, $"U{row}");
        if (rightCircuit > 0)
        {
          slots.Add(new PanelCircuitSlot
          {
            CircuitNumber = rightCircuit,
            Row = row,
            ConnectedKvaAddress = $"M{row}",
            LoadDescriptionAddress = $"N{row}",
            PolesAddress = $"Q{row}",
            BreakerAmpsAddress = $"R{row}",
            LoadTypeAddress = $"S{row}",
          });
        }
      }

      if (slots.Count == 0)
      {
        throw new InvalidOperationException(
          "The selected panel worksheet does not contain numbered circuits.");
      }
      return slots;
    }

    private static void ApplyCapacityAndSpareRules(
      Excel.Worksheet worksheet,
      List<PanelCircuitSlot> slots,
      int circuitCapacity)
    {
      if (circuitCapacity < 6 || circuitCapacity % 2 != 0)
      {
        throw new InvalidOperationException(
          "Panel circuit capacity must be an even number of at least six.");
      }

      int maximumCircuit = 0;
      List<int> conflicts = new List<int>();
      foreach (PanelCircuitSlot slot in slots)
      {
        maximumCircuit = Math.Max(maximumCircuit, slot.CircuitNumber);
        string description = ReadText(
          worksheet,
          slot.LoadDescriptionAddress);
        if ((slot.CircuitNumber > circuitCapacity ||
             IsReservedCircuit(slot.CircuitNumber, circuitCapacity)) &&
            !IsAvailableDescription(description))
        {
          conflicts.Add(slot.CircuitNumber);
        }
      }

      if (maximumCircuit < circuitCapacity)
      {
        throw new InvalidOperationException(
          $"The selected worksheet only contains circuits through " +
          $"{maximumCircuit}, not {circuitCapacity}.");
      }
      if (conflicts.Count > 0)
      {
        throw new InvalidOperationException(
          "Existing loads occupy circuits that must be reserved or hidden: " +
          string.Join(", ", conflicts) +
          ". Move those loads before applying this panel configuration.");
      }

      HashSet<int> processedRows = new HashSet<int>();
      foreach (PanelCircuitSlot slot in slots)
      {
        if (processedRows.Add(slot.Row))
        {
          SetRowHidden(
            worksheet,
            slot.Row,
            slot.CircuitNumber > circuitCapacity);
        }

        if (slot.CircuitNumber <= circuitCapacity &&
            IsReservedCircuit(slot.CircuitNumber, circuitCapacity))
        {
          WriteCell(worksheet, slot.LoadDescriptionAddress, "SPARE");
          WriteCell(worksheet, slot.LoadTypeAddress, null);
          WriteCell(worksheet, slot.PolesAddress, 1);
          WriteCell(worksheet, slot.BreakerAmpsAddress, 20);
          WriteCell(worksheet, slot.ConnectedKvaAddress, 0.0);
        }
      }
    }

    private static PanelCircuitSlot FindAvailableSlot(
      Excel.Worksheet worksheet,
      List<PanelCircuitSlot> slots,
      int circuitCapacity)
    {
      for (int parity = 1; parity >= 0; parity--)
      {
        PanelCircuitSlot best = null;
        foreach (PanelCircuitSlot slot in slots)
        {
          if (slot.CircuitNumber > circuitCapacity ||
              slot.CircuitNumber % 2 != parity ||
              IsReservedCircuit(slot.CircuitNumber, circuitCapacity) ||
              !IsAvailableDescription(
                ReadText(worksheet, slot.LoadDescriptionAddress)))
          {
            continue;
          }

          if (best == null || slot.CircuitNumber < best.CircuitNumber)
          {
            best = slot;
          }
        }
        if (best != null)
        {
          return best;
        }
      }
      return null;
    }

    private static bool IsReservedCircuit(
      int circuitNumber,
      int circuitCapacity)
    {
      return circuitNumber >
        circuitCapacity - GeneralCommands.ReservedPanelCircuitCount &&
        circuitNumber <= circuitCapacity;
    }

    private static bool IsAvailableDescription(string description)
    {
      string normalized = (description ?? string.Empty).Trim().ToUpperInvariant();
      return normalized.Length == 0 ||
        normalized.Contains("SPARE") ||
        normalized.Contains("SPACE");
    }

    private static string ReadText(
      Excel.Worksheet worksheet,
      string address)
    {
      object value = ReadCell(worksheet, address);
      return Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ??
        string.Empty;
    }

    private static int ReadInteger(
      Excel.Worksheet worksheet,
      string address)
    {
      object value = ReadCell(worksheet, address);
      if (value == null)
      {
        return 0;
      }
      try
      {
        return Convert.ToInt32(
          Math.Round(Convert.ToDouble(value, CultureInfo.InvariantCulture)),
          CultureInfo.InvariantCulture);
      }
      catch
      {
        Match match = Regex.Match(
          Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
          @"\d+");
        return match.Success
          ? int.Parse(match.Value, CultureInfo.InvariantCulture)
          : 0;
      }
    }

    private static object ReadCell(
      Excel.Worksheet worksheet,
      string address)
    {
      Excel.Range range = null;
      try
      {
        range = worksheet.Range[address];
        return range.Value2;
      }
      finally
      {
        ReleaseComObject(range);
      }
    }

    private static void WriteCell(
      Excel.Worksheet worksheet,
      string address,
      object value)
    {
      Excel.Range range = null;
      try
      {
        range = worksheet.Range[address];
        range.Value2 = value;
      }
      finally
      {
        ReleaseComObject(range);
      }
    }

    private static void SetRowHidden(
      Excel.Worksheet worksheet,
      int row,
      bool hidden)
    {
      Excel.Range rowRange = null;
      try
      {
        rowRange = worksheet.Range[$"{row}:{row}"];
        rowRange.EntireRow.Hidden = hidden;
      }
      finally
      {
        ReleaseComObject(rowRange);
      }
    }

    private static void ReleaseComObject(object value)
    {
      if (value == null || !Marshal.IsComObject(value))
      {
        return;
      }
      try
      {
        Marshal.FinalReleaseComObject(value);
      }
      catch
      {
      }
    }

    private sealed class PanelCircuitSlot
    {
      internal int CircuitNumber { get; set; }
      internal int Row { get; set; }
      internal string LoadTypeAddress { get; set; } = string.Empty;
      internal string PolesAddress { get; set; } = string.Empty;
      internal string BreakerAmpsAddress { get; set; } = string.Empty;
      internal string LoadDescriptionAddress { get; set; } = string.Empty;
      internal string ConnectedKvaAddress { get; set; } = string.Empty;
    }
  }
}
