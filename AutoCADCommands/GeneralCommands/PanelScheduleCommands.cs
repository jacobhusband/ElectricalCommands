using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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
    internal const int DefaultPanelCircuitCapacity = 42;
    internal const int DefaultPanelSpareCount = 6;

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

      string drawingDirectory = ResolveDrawingDirectory(document);
      if (existingSetting != null &&
          File.Exists(existingSetting.WorkbookPath))
      {
        dialog.FileName = existingSetting.WorkbookPath;
        dialog.InitialDirectory = Path.GetDirectoryName(
          existingSetting.WorkbookPath);
      }
      else
      {
        string suggestedWorkbook = FindSuggestedPanelScheduleWorkbook(
          drawingDirectory);
        if (suggestedWorkbook.Length > 0)
        {
          dialog.FileName = suggestedWorkbook;
          dialog.InitialDirectory = drawingDirectory;
          editor.WriteMessage(
            $"\nSuggested panel schedule found next to the drawing: " +
            $"\"{suggestedWorkbook}\".");
        }
        else if (drawingDirectory.Length > 0)
        {
          dialog.InitialDirectory = drawingDirectory;
        }
      }

      if (dialog.ShowDialog() != true)
      {
        editor.WriteMessage("\nPanel schedule selection canceled.");
        return;
      }

      int spareCount = 0;

      try
      {
        string worksheetName = PanelScheduleWorkbookAllocator.PreparePanel(
          dialog.FileName,
          panelName,
          circuitCapacity: 0,
          spareCount: spareCount,
          out int detectedCapacity,
          out PanelCircuitCounts counts);
        ElectricalDrawingSettingsStore.WritePanelSchedule(
          document.Database,
          Path.GetFullPath(dialog.FileName),
          detectedCapacity,
          spareCount);
        HomerunSettingsPalette.Refresh();

        editor.WriteMessage(
          $"\nPanel schedule linked to worksheet \"{worksheetName}\". Autodetected {detectedCapacity} active circuits (1-{detectedCapacity})." +
          $"\n{FormatPanelCircuitStatus(panelName, counts)}");
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to link the panel schedule: {ex.Message}");
      }
    }

    [CommandMethod("SETSPARES", CommandFlags.Modal)]
    [CommandMethod("SPSPARES", CommandFlags.Modal)]
    [CommandMethod("SETSPANELSPARES", CommandFlags.Modal)]
    public static void SetPanelSparesCommand()
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
          "\nSet the panel name with SETPANELNAME (SPN) first.");
        return;
      }

      if (!ElectricalDrawingSettingsStore.TryReadPanelSchedule(
        document.Database,
        out var panelSchedule) ||
        string.IsNullOrWhiteSpace(panelSchedule.WorkbookPath) ||
        !File.Exists(panelSchedule.WorkbookPath))
      {
        editor.WriteMessage(
          "\nNo linked panel schedule found. Run SETPANELSCHEDULE (SPS) first.");
        return;
      }

      int defaultSpares = Math.Max(
        0,
        Math.Min(panelSchedule.SpareCount, panelSchedule.CircuitCapacity));
      PromptIntegerOptions spareOptions = new PromptIntegerOptions(
        $"\nEnter the number of spare circuits to reserve in panel {panelName} " +
        $"<{defaultSpares}>: ")
      {
        AllowNone = true,
        AllowNegative = false,
        AllowZero = true,
        DefaultValue = defaultSpares,
        LowerLimit = 0,
        UpperLimit = panelSchedule.CircuitCapacity,
        UseDefaultValue = true,
      };
      PromptIntegerResult spareResult = editor.GetInteger(spareOptions);
      if (spareResult.Status != PromptStatus.OK &&
          spareResult.Status != PromptStatus.None)
      {
        editor.WriteMessage("\nSet panel spares canceled.");
        return;
      }

      int spareCount = spareResult.Status == PromptStatus.None
        ? defaultSpares
        : spareResult.Value;

      try
      {
        string worksheetName = PanelScheduleWorkbookAllocator.PreparePanel(
          panelSchedule.WorkbookPath,
          panelName,
          panelSchedule.CircuitCapacity,
          spareCount,
          out PanelCircuitCounts counts);

        ElectricalDrawingSettingsStore.WritePanelSchedule(
          document.Database,
          panelSchedule.WorkbookPath,
          panelSchedule.CircuitCapacity,
          spareCount);
        HomerunSettingsPalette.Refresh();

        editor.WriteMessage(
          $"\nPanel {panelName} spares set to {spareCount} on worksheet \"{worksheetName}\"." +
          $"\n{FormatPanelCircuitStatus(panelName, counts)}");
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to update panel spares: {ex.Message}");
      }
    }

    private static string ResolveDrawingDirectory(Document document)
    {
      try
      {
        string drawingPath = document?.Database?.Filename ?? string.Empty;
        string directory = Path.GetDirectoryName(drawingPath);
        return !string.IsNullOrWhiteSpace(directory) &&
          Directory.Exists(directory)
          ? directory
          : string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }

    private static string FindSuggestedPanelScheduleWorkbook(
      string drawingDirectory)
    {
      if (string.IsNullOrWhiteSpace(drawingDirectory) ||
          !Directory.Exists(drawingDirectory))
      {
        return string.Empty;
      }

      try
      {
        List<string> candidates = new List<string>();
        foreach (string path in Directory.GetFiles(drawingDirectory))
        {
          string fileName = Path.GetFileName(path);
          string extension = Path.GetExtension(path);
          if (fileName.StartsWith("~$", StringComparison.OrdinalIgnoreCase) ||
              (!string.Equals(
                 extension,
                 ".xls",
                 StringComparison.OrdinalIgnoreCase) &&
               !string.Equals(
                 extension,
                 ".xlsx",
                 StringComparison.OrdinalIgnoreCase)))
          {
            continue;
          }
          candidates.Add(path);
        }

        candidates.Sort(StringComparer.OrdinalIgnoreCase);
        foreach (string candidate in candidates)
        {
          if (Path.GetFileNameWithoutExtension(candidate).IndexOf(
            "panel",
            StringComparison.OrdinalIgnoreCase) >= 0)
          {
            return candidate;
          }
        }
        return candidates.Count > 0 ? candidates[0] : string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }

    private static string BuildReservedCircuitSummary(
      int circuitCapacity,
      int spareCount)
    {
      if (spareCount <= 0)
      {
        return "no circuits";
      }
      List<string> numbers = new List<string>();
      for (int circuit = Math.Max(1, circuitCapacity - spareCount + 1);
           circuit <= circuitCapacity;
           circuit++)
      {
        numbers.Add(circuit.ToString(CultureInfo.InvariantCulture));
      }
      return "circuits " + string.Join(", ", numbers);
    }

    internal static string FormatPanelCircuitStatus(
      string panelName,
      PanelCircuitCounts counts)
    {
      if (counts == null)
      {
        return string.Empty;
      }
      return $"Panel {panelName} remaining: {counts.Spares} spare(s), " +
        $"{counts.Spaces} space(s), {counts.Empties} empty circuit(s) " +
        $"({counts.AvailableTotal} total available).";
    }
  }

  internal sealed class PanelCircuitCounts
  {
    internal int Spares { get; set; }
    internal int Spaces { get; set; }
    internal int Empties { get; set; }
    internal int ActiveLoads { get; set; }
    internal int TotalCircuits { get; set; }
    internal int AvailableTotal => Spares + Spaces + Empties;
  }

  internal sealed class PanelScheduleAllocationResult
  {
    internal int CircuitNumber { get; set; }
    internal string CircuitLabel { get; set; } = string.Empty;
    internal string WorksheetName { get; set; } = string.Empty;
    internal double ConnectedWatts { get; set; }
    internal PanelCircuitCounts RemainingCounts { get; set; }
  }

  internal sealed class PanelScheduleCircuitRequest
  {
    internal double ConnectedWatts { get; set; }
    internal string LoadDescription { get; set; } = "RECEPTACLES";
    internal string LoadTypeCode { get; set; } = "G";
    internal int Poles { get; set; } = 1;
    internal int BreakerAmps { get; set; } = 20;
  }

  internal static class PanelScheduleWorkbookAllocator
  {
    private const int AciesStartRow = 7;
    private const int AciesMaxRow = 48;
    private const int GeneratedStartRow = 8;
    private const int GeneratedMaxRow = 28;
    private const int BlackFontColor = 0x000000;
    private const int ExistingCircuitFontColor = 0xA6A6A6;

    internal static string PreparePanel(
      string workbookPath,
      string panelName,
      int circuitCapacity,
      int spareCount,
      out int resolvedCapacity,
      out PanelCircuitCounts counts)
    {
      PanelCircuitCounts localCounts = null;
      int localCapacity = circuitCapacity;
      string sheetName = ExecuteWithPanelWorksheet(
        workbookPath,
        panelName,
        worksheet =>
        {
          if (localCapacity <= 0)
          {
            localCapacity = DetectNonHiddenCircuitCapacity(worksheet);
          }
          List<PanelCircuitSlot> slots = BuildCircuitSlots(worksheet);
          ApplyCapacityAndSpareRules(worksheet, slots, localCapacity, spareCount);
          localCounts = CountRemainingSlots(worksheet, slots, localCapacity);
          return worksheet.Name;
        });
      resolvedCapacity = localCapacity;
      counts = localCounts;
      return sheetName;
    }

    internal static string PreparePanel(
      string workbookPath,
      string panelName,
      int circuitCapacity,
      int spareCount,
      out PanelCircuitCounts counts)
    {
      return PreparePanel(
        workbookPath,
        panelName,
        circuitCapacity,
        spareCount,
        out _,
        out counts);
    }

    internal static List<PanelScheduleAllocationResult>
      AllocateReceptacleCircuits(
        string workbookPath,
        string panelName,
        int circuitCapacity,
        int spareCount,
        IList<PanelScheduleCircuitRequest> requests)
    {
      if (requests == null || requests.Count == 0)
      {
        throw new InvalidOperationException(
          "At least one receptacle circuit is required.");
      }

      foreach (PanelScheduleCircuitRequest request in requests)
      {
        if (request == null ||
            request.ConnectedWatts <= 0.0 ||
            double.IsNaN(request.ConnectedWatts) ||
            double.IsInfinity(request.ConnectedWatts) ||
            request.Poles < 1 ||
            request.Poles > 3 ||
            request.BreakerAmps <= 0)
        {
          throw new InvalidOperationException(
            "Each circuit requires a positive load and breaker rating with " +
            "one, two, or three poles.");
        }
      }

      return ExecuteWithPanelWorksheet(
        workbookPath,
        panelName,
        worksheet =>
        {
          List<PanelCircuitSlot> slots = BuildCircuitSlots(worksheet);
          ApplyCapacityAndSpareRules(worksheet, slots, circuitCapacity, spareCount);
          bool isExistingPanel = IsExistingPanel(worksheet, panelName);

          List<PanelScheduleAllocationResult> results =
            new List<PanelScheduleAllocationResult>();
          foreach (PanelScheduleCircuitRequest request in requests)
          {
            List<PanelCircuitSlot> available = FindAvailableSlots(
              worksheet,
              slots,
              circuitCapacity,
              spareCount,
              request.Poles);
            if (available.Count != request.Poles)
            {
              string spareDetail = spareCount > 0
                ? $" (excluding {spareCount} reserved spares)"
                : string.Empty;
              throw new InvalidOperationException(
                $"Panel {panelName} does not have enough usable SPARE, " +
                $"SPACE, or empty consecutive circuit positions{spareDetail} for a {request.Poles}-pole load.");
            }

            bool[] hadExistingBreakers = available
              .Select(slot =>
                isExistingPanel && HasExistingCircuitBreaker(
                  worksheet,
                  slot))
              .ToArray();

            string loadDescription =
              (request.LoadDescription ?? string.Empty).Trim();
            if (loadDescription.Length == 0)
            {
              loadDescription = "RECEPTACLES";
            }

            string loadTypeCode =
              (request.LoadTypeCode ?? string.Empty).Trim().ToUpperInvariant();
            if (loadTypeCode.Length == 0)
            {
              loadTypeCode = "G";
            }
            double connectedKvaPerPole = Math.Round(
              request.ConnectedWatts / 1000.0 / request.Poles,
              3);
            for (int poleIndex = 0;
                 poleIndex < available.Count;
                 poleIndex++)
            {
              PanelCircuitSlot slot = available[poleIndex];
              bool isPrimary = poleIndex == 0;
              WriteCell(
                worksheet,
                slot.LoadDescriptionAddress,
                isPrimary ? loadDescription : null);
              WriteCell(worksheet, slot.LoadTypeAddress, loadTypeCode);
              WriteCell(
                worksheet,
                slot.NotesAddress,
                null);
              WriteCell(
                worksheet,
                slot.PolesAddress,
                isPrimary ? (object)request.Poles : null);
              WriteCell(
                worksheet,
                slot.BreakerAmpsAddress,
                isPrimary ? (object)request.BreakerAmps : null);
              WriteCell(
                worksheet,
                slot.ConnectedKvaAddress,
                connectedKvaPerPole);
              ApplyCircuitEntryFont(
                worksheet,
                slot,
                isExistingPanel,
                hadExistingBreakers[poleIndex]);
            }

            string circuitLabel = string.Join(
              "/",
              available.Select(
                slot => slot.CircuitNumber.ToString(
                  CultureInfo.InvariantCulture)));

            results.Add(new PanelScheduleAllocationResult
            {
              CircuitNumber = available[0].CircuitNumber,
              CircuitLabel = circuitLabel,
              WorksheetName = worksheet.Name,
              ConnectedWatts = request.ConnectedWatts,
            });
          }

          PanelCircuitCounts remainingCounts =
            CountRemainingSlots(worksheet, slots, circuitCapacity);
          foreach (PanelScheduleAllocationResult result in results)
          {
            result.RemainingCounts = remainingCounts;
          }

          return results;
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

    private static bool IsExistingPanel(
      Excel.Worksheet worksheet,
      string panelName)
    {
      string[] candidates =
      {
        panelName,
        worksheet.Name,
        ReadText(worksheet, "A1"),
        ReadText(worksheet, "A2"),
        ReadText(worksheet, "A3"),
        ReadText(worksheet, "A4"),
      };
      foreach (string candidate in candidates)
      {
        if (Regex.IsMatch(
          candidate ?? string.Empty,
          @"(?i)\(\s*(?:E|EX|EXISTING|RL|RELOCATED)\s*\)"))
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
        text = quoted.Groups[1].Value.Trim();
      }
      else
      {
        text = Regex.Replace(text, @"^\s*\([^)]*\)\s*", string.Empty);
        text = Regex.Replace(
          text,
          @"^\s*(?:PANEL|PNL)\s*",
          string.Empty,
          RegexOptions.IgnoreCase);
      }

      text = Regex.Replace(
        text,
        @"\(\s*(?:E|EX|EXISTING|RL|RELOCATED)\s*\)",
        " ",
        RegexOptions.IgnoreCase);
      text = Regex.Replace(text, @"\s+", " ");
      return text.Trim(' ', '\"', '\'').ToUpperInvariant();
    }

    internal static int DetectNonHiddenCircuitCapacity(
      Excel.Worksheet worksheet)
    {
      if (!TryFindCircuitStartRow(worksheet, out int startRow))
      {
        return GeneralCommands.DefaultPanelCircuitCapacity;
      }

      string rightCktCol = DetectRightCircuitColumn(worksheet, startRow);

      int highestCircuit = 0;
      for (int row = startRow; row < startRow + 42; row++)
      {
        int leftCircuit = ReadInteger(worksheet, $"A{row}");
        int rightCircuit = ReadInteger(worksheet, $"{rightCktCol}{row}");
        bool hasLeftCircuit = IsValidCircuitNumber(leftCircuit);
        bool hasRightCircuit = IsValidCircuitNumber(rightCircuit);
        if (!hasLeftCircuit && !hasRightCircuit)
        {
          // The first unnumbered pair marks the end of the circuit table.
          // Do not interpret demand-calculation or sub-panel summary rows as
          // additional circuit positions.
          break;
        }

        Excel.Range rowRange = null;
        try
        {
          rowRange = worksheet.Rows[row] as Excel.Range;
          if (rowRange != null && Convert.ToBoolean(rowRange.Hidden))
          {
            // Row is hidden, ignore
            continue;
          }
        }
        finally
        {
          ReleaseComObject(rowRange);
        }

        int inferredLeftCircuit = hasLeftCircuit
          ? leftCircuit
          : rightCircuit - 1;
        int inferredRightCircuit = hasRightCircuit
          ? rightCircuit
          : leftCircuit + 1;
        highestCircuit = Math.Max(
          highestCircuit,
          Math.Max(inferredLeftCircuit, inferredRightCircuit));
      }

      if (highestCircuit > 0)
      {
        if (highestCircuit % 2 != 0)
        {
          highestCircuit++;
        }
        return highestCircuit;
      }

      return GeneralCommands.DefaultPanelCircuitCapacity;
    }

    private static List<PanelCircuitSlot> BuildCircuitSlots(
      Excel.Worksheet worksheet)
    {
      bool hasExplicitCircuitNumbers = TryFindCircuitStartRow(
        worksheet,
        out int startRow);
      if (!hasExplicitCircuitNumbers)
      {
        startRow = GeneratedStartRow;
      }

      string rightCktCol = DetectRightCircuitColumn(worksheet, startRow);
      int endRow = hasExplicitCircuitNumbers
        ? startRow + 41
        : GeneratedMaxRow;

      List<PanelCircuitSlot> slots = new List<PanelCircuitSlot>();
      for (int row = startRow; row <= endRow; row++)
      {
        int leftCircuit = ReadInteger(worksheet, $"A{row}");
        int oddCircuit = (row - startRow) * 2 + 1;
        int rightCircuit = ReadInteger(worksheet, $"{rightCktCol}{row}");

        if (hasExplicitCircuitNumbers)
        {
          bool hasLeftCircuit = IsValidCircuitNumber(leftCircuit);
          bool hasRightCircuit = IsValidCircuitNumber(rightCircuit);
          if (!hasLeftCircuit && !hasRightCircuit)
          {
            break;
          }

          leftCircuit = hasLeftCircuit
            ? leftCircuit
            : rightCircuit - 1;
          rightCircuit = hasRightCircuit
            ? rightCircuit
            : leftCircuit + 1;
        }
        else
        {
          // Preserve the original generated-workbook fallback, whose circuit
          // rows are fixed at 8-28 even when the number cells are blank.
          leftCircuit = oddCircuit;
          rightCircuit = oddCircuit + 1;
        }

        if (rightCktCol == "U")
        {
          slots.Add(new PanelCircuitSlot
          {
            CircuitNumber = leftCircuit,
            Row = row,
            LoadTypeAddress = $"C{row}",
            NotesAddress = $"B{row}",
            PolesAddress = $"D{row}",
            BreakerAmpsAddress = $"E{row}",
            LoadDescriptionAddress = $"F{row}",
            ConnectedKvaAddress = $"K{row}",
          });
          slots.Add(new PanelCircuitSlot
          {
            CircuitNumber = rightCircuit,
            Row = row,
            ConnectedKvaAddress = $"M{row}",
            LoadDescriptionAddress = $"N{row}",
            NotesAddress = $"T{row}",
            PolesAddress = $"Q{row}",
            BreakerAmpsAddress = $"R{row}",
            LoadTypeAddress = $"S{row}",
          });
        }
        else
        {
          // Standard ACIES / Charles Schwab / BofA 19-column layout (Right CKT Col = S, L:N merged for description)
          slots.Add(new PanelCircuitSlot
          {
            CircuitNumber = leftCircuit,
            Row = row,
            LoadTypeAddress = $"C{row}",
            NotesAddress = $"B{row}",
            PolesAddress = $"D{row}",
            BreakerAmpsAddress = $"E{row}",
            LoadDescriptionAddress = $"F{row}",
            ConnectedKvaAddress = $"H{row}",
          });
          slots.Add(new PanelCircuitSlot
          {
            CircuitNumber = rightCircuit,
            Row = row,
            ConnectedKvaAddress = $"K{row}",
            LoadDescriptionAddress = $"L{row}",
            NotesAddress = $"R{row}",
            PolesAddress = $"O{row}",
            BreakerAmpsAddress = $"P{row}",
            LoadTypeAddress = $"Q{row}",
          });
        }

        if (leftCircuit >= 84 || rightCircuit >= 84)
        {
          break;
        }
      }

      if (slots.Count == 0)
      {
        throw new InvalidOperationException(
          "The selected panel worksheet does not contain numbered circuits.");
      }
      return slots;
    }

    private static bool TryFindCircuitStartRow(
      Excel.Worksheet worksheet,
      out int startRow)
    {
      for (int row = 5; row <= 15; row++)
      {
        if (ReadInteger(worksheet, $"A{row}") == 1)
        {
          startRow = row;
          return true;
        }
      }

      startRow = -1;
      return false;
    }

    private static string DetectRightCircuitColumn(
      Excel.Worksheet worksheet,
      int startRow)
    {
      if (ReadInteger(worksheet, $"U{startRow}") == 2)
      {
        return "U";
      }
      if (ReadInteger(worksheet, $"S{startRow}") == 2)
      {
        return "S";
      }
      if (ReadInteger(worksheet, $"T{startRow}") == 2)
      {
        return "T";
      }
      if (string.Equals(
            ReadText(worksheet, "U5"),
            "CKT#",
            StringComparison.OrdinalIgnoreCase) ||
          string.Equals(
            ReadText(worksheet, "U6"),
            "CKT#",
            StringComparison.OrdinalIgnoreCase))
      {
        return "U";
      }
      return "S";
    }

    private static bool IsValidCircuitNumber(int circuitNumber)
    {
      return circuitNumber > 0 && circuitNumber <= 84;
    }

    private static void ApplyCapacityAndSpareRules(
      Excel.Worksheet worksheet,
      List<PanelCircuitSlot> slots,
      int circuitCapacity,
      int spareCount)
    {
      if (circuitCapacity < 2 || circuitCapacity % 2 != 0)
      {
        throw new InvalidOperationException(
          "Panel circuit capacity must be an even number of at least two.");
      }

      int maximumCircuit = 0;
      List<int> hiddenConflicts = new List<int>();
      foreach (PanelCircuitSlot slot in slots)
      {
        maximumCircuit = Math.Max(maximumCircuit, slot.CircuitNumber);
        if (slot.CircuitNumber > circuitCapacity &&
            !IsAvailableSlot(worksheet, slot))
        {
          hiddenConflicts.Add(slot.CircuitNumber);
        }
      }

      if (maximumCircuit < circuitCapacity)
      {
        throw new InvalidOperationException(
          $"The selected worksheet only contains circuits through " +
          $"{maximumCircuit}, not {circuitCapacity}.");
      }
      if (hiddenConflicts.Count > 0)
      {
        throw new InvalidOperationException(
          "Existing loads occupy circuits that would be hidden by the selected capacity: " +
          string.Join(", ", hiddenConflicts) +
          ". Move those loads before reducing the panel circuit count.");
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
      }
    }

    private static List<PanelCircuitSlot> FindAvailableSlots(
      Excel.Worksheet worksheet,
      List<PanelCircuitSlot> slots,
      int circuitCapacity,
      int spareCount,
      int poles)
    {
      var activeSlots = slots
        .Where(slot => slot.CircuitNumber <= circuitCapacity)
        .ToList();

      // Count total spares currently in the panel
      int totalSparesInPanel = activeSlots.Count(s =>
      {
        string desc = ReadSlotDescription(worksheet, s).ToUpperInvariant();
        return desc.Contains("SPARE");
      });

      // Allowable spares that can be converted to new circuits
      int allowableSparesToUse = Math.Max(0, totalSparesInPanel - spareCount);

      Dictionary<int, PanelCircuitSlot> slotsByCircuit = activeSlots.ToDictionary(
        slot => slot.CircuitNumber);

      // Pass 1: Try to find consecutive EMPTY or SPACE slots (without consuming any spares)
      for (int parity = 1; parity >= 0; parity--)
      {
        foreach (PanelCircuitSlot startingSlot in activeSlots
          .Where(slot => slot.CircuitNumber % 2 == parity)
          .OrderBy(slot => slot.CircuitNumber))
        {
          var group = new List<PanelCircuitSlot>();
          for (int poleIndex = 0; poleIndex < poles; poleIndex++)
          {
            int circuitNumber = startingSlot.CircuitNumber + poleIndex * 2;
            if (!slotsByCircuit.TryGetValue(circuitNumber, out PanelCircuitSlot slot) ||
                !IsAvailableSlot(worksheet, slot))
            {
              group.Clear();
              break;
            }

            string desc = ReadSlotDescription(worksheet, slot).ToUpperInvariant();
            if (desc.Contains("SPARE"))
            {
              // Do not consume spares in pass 1
              group.Clear();
              break;
            }
            group.Add(slot);
          }
          if (group.Count == poles)
          {
            return group;
          }
        }
      }

      // Pass 2: If allowableSparesToUse >= poles, find consecutive available slots that may include SPARE
      if (allowableSparesToUse >= poles)
      {
        for (int parity = 1; parity >= 0; parity--)
        {
          foreach (PanelCircuitSlot startingSlot in activeSlots
            .Where(slot => slot.CircuitNumber % 2 == parity)
            .OrderBy(slot => slot.CircuitNumber))
          {
            var group = new List<PanelCircuitSlot>();
            int sparesInGroup = 0;
            for (int poleIndex = 0; poleIndex < poles; poleIndex++)
            {
              int circuitNumber = startingSlot.CircuitNumber + poleIndex * 2;
              if (!slotsByCircuit.TryGetValue(circuitNumber, out PanelCircuitSlot slot) ||
                  !IsAvailableSlot(worksheet, slot))
              {
                group.Clear();
                break;
              }

              string desc = ReadSlotDescription(worksheet, slot).ToUpperInvariant();
              if (desc.Contains("SPARE"))
              {
                sparesInGroup++;
              }
              group.Add(slot);
            }
            if (group.Count == poles && sparesInGroup <= allowableSparesToUse)
            {
              return group;
            }
          }
        }
      }

      return new List<PanelCircuitSlot>();
    }

    private static PanelCircuitCounts CountRemainingSlots(
      Excel.Worksheet worksheet,
      List<PanelCircuitSlot> slots,
      int circuitCapacity)
    {
      var counts = new PanelCircuitCounts
      {
        TotalCircuits = circuitCapacity,
      };

      foreach (PanelCircuitSlot slot in slots.Where(s => s.CircuitNumber <= circuitCapacity))
      {
        string description = ReadSlotDescription(worksheet, slot).Trim().ToUpperInvariant();
        string loadType = ReadText(worksheet, slot.LoadTypeAddress).Trim();
        string kvaText = ReadText(worksheet, slot.ConnectedKvaAddress).Trim();
        int breaker = ReadInteger(worksheet, slot.BreakerAmpsAddress);

        if (description.Contains("SPARE"))
        {
          counts.Spares++;
        }
        else if (description.Contains("SPACE"))
        {
          counts.Spaces++;
        }
        else if ((description.Length == 0 || description == "---" || description == "-") &&
                 string.IsNullOrWhiteSpace(loadType) &&
                 breaker == 0 &&
                 (string.IsNullOrWhiteSpace(kvaText) || kvaText == "0" || kvaText == "0.00" || kvaText == "0.0"))
        {
          counts.Empties++;
        }
        else
        {
          counts.ActiveLoads++;
        }
      }

      return counts;
    }

    private static string ReadSlotDescription(
      Excel.Worksheet worksheet,
      PanelCircuitSlot slot)
    {
      string text = ReadText(worksheet, slot.LoadDescriptionAddress);
      if (!string.IsNullOrWhiteSpace(text) && text != "---" && text != "-")
      {
        return text;
      }
      if (slot.LoadDescriptionAddress.StartsWith("F", StringComparison.OrdinalIgnoreCase))
      {
        string alt = ReadText(worksheet, $"G{slot.Row}");
        if (!string.IsNullOrWhiteSpace(alt) && alt != "---" && alt != "-")
        {
          return alt;
        }
      }
      else if (slot.LoadDescriptionAddress.StartsWith("L", StringComparison.OrdinalIgnoreCase) ||
               slot.LoadDescriptionAddress.StartsWith("N", StringComparison.OrdinalIgnoreCase))
      {
        string altN = ReadText(worksheet, $"N{slot.Row}");
        if (!string.IsNullOrWhiteSpace(altN) && altN != "---" && altN != "-")
        {
          return altN;
        }
        string altM = ReadText(worksheet, $"M{slot.Row}");
        if (!string.IsNullOrWhiteSpace(altM) && altM != "---" && altM != "-")
        {
          return altM;
        }
        string altL = ReadText(worksheet, $"L{slot.Row}");
        if (!string.IsNullOrWhiteSpace(altL) && altL != "---" && altL != "-")
        {
          return altL;
        }
      }
      return text;
    }

    private static bool IsAvailableSlot(
      Excel.Worksheet worksheet,
      PanelCircuitSlot slot)
    {
      string description = ReadSlotDescription(
        worksheet,
        slot);
      string normalized = description.Trim().ToUpperInvariant();
      if (normalized.Contains("SPARE") || normalized.Contains("SPACE"))
      {
        return true;
      }
      if (normalized.Length == 0 || normalized == "---" || normalized == "-")
      {
        string loadType = ReadText(worksheet, slot.LoadTypeAddress);
        int breaker = ReadInteger(worksheet, slot.BreakerAmpsAddress);
        string kva = ReadText(worksheet, slot.ConnectedKvaAddress);
        return string.IsNullOrWhiteSpace(loadType) &&
               breaker == 0 &&
               (string.IsNullOrWhiteSpace(kva) || kva == "0" || kva == "0.00" || kva == "0.0");
      }
      return false;
    }

    private static bool HasExistingCircuitBreaker(
      Excel.Worksheet worksheet,
      PanelCircuitSlot slot)
    {
      string description = ReadSlotDescription(
        worksheet,
        slot).ToUpperInvariant();
      if (description.Contains("SPACE"))
      {
        return false;
      }

      return description.Contains("SPARE") ||
        ReadInteger(worksheet, slot.BreakerAmpsAddress) > 0;
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
        if (Convert.ToBoolean(range.MergeCells))
        {
          Excel.Range mergeArea = null;
          try
          {
            mergeArea = range.MergeArea;
            // Check top-left cell first
            Excel.Range firstCell = mergeArea.Cells[1, 1] as Excel.Range;
            try
            {
              object firstVal = firstCell?.Value2;
              if (firstVal != null && !string.IsNullOrWhiteSpace(Convert.ToString(firstVal, CultureInfo.InvariantCulture)))
              {
                return firstVal;
              }
            }
            finally
            {
              ReleaseComObject(firstCell);
            }

            // If top-left cell is empty, check any other cell in the merge area
            foreach (Excel.Range cell in mergeArea.Cells)
            {
              try
              {
                object cellVal = cell.Value2;
                if (cellVal != null && !string.IsNullOrWhiteSpace(Convert.ToString(cellVal, CultureInfo.InvariantCulture)))
                {
                  return cellVal;
                }
              }
              finally
              {
                ReleaseComObject(cell);
              }
            }
          }
          finally
          {
            ReleaseComObject(mergeArea);
          }
        }
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
        if (Convert.ToBoolean(range.MergeCells))
        {
          Excel.Range mergeArea = null;
          Excel.Range firstCell = null;
          try
          {
            mergeArea = range.MergeArea;
            firstCell = mergeArea.Cells[1, 1] as Excel.Range;
            if (firstCell != null)
            {
              firstCell.Value2 = value;
              return;
            }
          }
          finally
          {
            ReleaseComObject(firstCell);
            ReleaseComObject(mergeArea);
          }
        }
        range.Value2 = value;
      }
      finally
      {
        ReleaseComObject(range);
      }
    }

    private static void ApplyCircuitEntryFont(
      Excel.Worksheet worksheet,
      PanelCircuitSlot slot,
      bool isExistingPanel,
      bool hadExistingBreaker)
    {
      bool emphasizeNewLoad = isExistingPanel;
      ApplyCellFont(
        worksheet,
        slot.LoadDescriptionAddress,
        emphasizeNewLoad,
        false,
        BlackFontColor);
      ApplyCellFont(
        worksheet,
        slot.ConnectedKvaAddress,
        emphasizeNewLoad,
        false,
        BlackFontColor);

      // Notes and load-type designations describe the new work and always use
      // the normal black schedule style, regardless of the panel status.
      ApplyCellFont(
        worksheet,
        slot.NotesAddress,
        false,
        false,
        BlackFontColor);
      ApplyCellFont(
        worksheet,
        slot.LoadTypeAddress,
        false,
        false,
        BlackFontColor);

      bool showAsExistingBreaker = isExistingPanel && hadExistingBreaker;
      bool emphasizeNewBreaker = isExistingPanel && !hadExistingBreaker;
      int breakerColor = showAsExistingBreaker
        ? ExistingCircuitFontColor
        : BlackFontColor;
      ApplyCellFont(
        worksheet,
        slot.PolesAddress,
        emphasizeNewBreaker,
        showAsExistingBreaker,
        breakerColor);
      ApplyCellFont(
        worksheet,
        slot.BreakerAmpsAddress,
        emphasizeNewBreaker,
        showAsExistingBreaker,
        breakerColor);
    }

    private static void ApplyCellFont(
      Excel.Worksheet worksheet,
      string address,
      bool bold,
      bool italic,
      int color)
    {
      Excel.Range range = null;
      Excel.Range targetRange = null;
      Excel.Font font = null;
      try
      {
        range = worksheet.Range[address];
        targetRange = Convert.ToBoolean(range.MergeCells) ? range.MergeArea : range;
        font = targetRange.Font;
        font.Bold = bold;
        font.Italic = italic;
        font.Color = color;
      }
      finally
      {
        ReleaseComObject(font);
        if (!ReferenceEquals(targetRange, range))
        {
          ReleaseComObject(targetRange);
        }
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
      internal string NotesAddress { get; set; } = string.Empty;
      internal string PolesAddress { get; set; } = string.Empty;
      internal string BreakerAmpsAddress { get; set; } = string.Empty;
      internal string LoadDescriptionAddress { get; set; } = string.Empty;
      internal string ConnectedKvaAddress { get; set; } = string.Empty;
    }
  }
}
