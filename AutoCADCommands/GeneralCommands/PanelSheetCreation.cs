using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.EditorInput;
using Excel = Microsoft.Office.Interop.Excel;

namespace ElectricalCommands
{
  internal enum PanelHeaderField
  {
    Voltage,
    BusRating,
    Main,
    Wires,
    Phase,
    Type,
    Mounting,
    AicRating,
    Condition,
  }

  /// <summary>The labeled fields in the header block of a panel worksheet.</summary>
  internal sealed class PanelHeaderValues
  {
    internal string Voltage { get; set; } = string.Empty;
    internal string BusRating { get; set; } = string.Empty;
    internal string Main { get; set; } = string.Empty;
    internal string Wires { get; set; } = string.Empty;
    internal string Phase { get; set; } = string.Empty;
    internal string Type { get; set; } = string.Empty;
    internal string Mounting { get; set; } = string.Empty;
    internal string AicRating { get; set; } = string.Empty;
    internal string Condition { get; set; } = string.Empty;

    internal string Get(PanelHeaderField field)
    {
      switch (field)
      {
        case PanelHeaderField.Voltage: return Voltage;
        case PanelHeaderField.BusRating: return BusRating;
        case PanelHeaderField.Main: return Main;
        case PanelHeaderField.Wires: return Wires;
        case PanelHeaderField.Phase: return Phase;
        case PanelHeaderField.Type: return Type;
        case PanelHeaderField.Mounting: return Mounting;
        case PanelHeaderField.AicRating: return AicRating;
        default: return Condition;
      }
    }

    internal void Set(PanelHeaderField field, string value)
    {
      value = value ?? string.Empty;
      switch (field)
      {
        case PanelHeaderField.Voltage: Voltage = value; break;
        case PanelHeaderField.BusRating: BusRating = value; break;
        case PanelHeaderField.Main: Main = value; break;
        case PanelHeaderField.Wires: Wires = value; break;
        case PanelHeaderField.Phase: Phase = value; break;
        case PanelHeaderField.Type: Type = value; break;
        case PanelHeaderField.Mounting: Mounting = value; break;
        case PanelHeaderField.AicRating: AicRating = value; break;
        default: Condition = value; break;
      }
    }
  }

  /// <summary>An existing panel worksheet that can be copied as a template.</summary>
  internal sealed class PanelTemplateInfo
  {
    internal string SheetName { get; set; } = string.Empty;
    internal int CircuitCapacity { get; set; }
    internal PanelHeaderValues Header { get; set; } = new PanelHeaderValues();

    internal string DisplayText =>
      $"{SheetName}  ({CircuitCapacity} circuits)";
  }

  internal sealed class PanelTemplateCatalog
  {
    internal List<PanelTemplateInfo> Templates { get; } =
      new List<PanelTemplateInfo>();

    /// <summary>Every worksheet name in the workbook, for uniqueness checks.</summary>
    internal List<string> AllSheetNames { get; } = new List<string>();
  }

  internal sealed class PanelSheetRequest
  {
    internal string TemplateSheetName { get; set; } = string.Empty;
    internal string NewSheetName { get; set; } = string.Empty;
    internal string PanelName { get; set; } = string.Empty;
    internal PanelHeaderValues Header { get; set; } = new PanelHeaderValues();
    internal int CircuitCount { get; set; }
    internal IReadOnlyList<string> BreakerPhases { get; set; } =
      new[] { "A", "B", "C" };
  }

  internal sealed class PanelSheetCreationResult
  {
    internal string SheetName { get; set; } = string.Empty;
    internal int CircuitCount { get; set; }

    /// <summary>Header fields the template has no labeled cell for.</summary>
    internal List<PanelHeaderField> MissingFields { get; } =
      new List<PanelHeaderField>();
  }

  internal sealed class PanelWorksheetNotFoundException
    : InvalidOperationException
  {
    internal string PanelName { get; }
    internal IReadOnlyList<string> AvailableWorksheets { get; }

    internal PanelWorksheetNotFoundException(
      string panelName,
      IReadOnlyList<string> availableWorksheets)
      : base(
        $"Could not find panel {panelName} in the workbook. Available panel " +
        $"worksheets: {(availableWorksheets.Count == 0 ? "none" : string.Join(", ", availableWorksheets))}. " +
        "Run SETPANELSCHEDULE (SPS) to create a worksheet for it.")
    {
      PanelName = panelName;
      AvailableWorksheets = availableWorksheets;
    }
  }

  /// <summary>
  /// Validation and formatting for new panel worksheets. Free of Excel and
  /// AutoCAD calls so it can be exercised on its own.
  /// </summary>
  internal static class PanelSheetRules
  {
    internal const int MaximumSheetNameLength = 31;
    internal const int MinimumCircuitCount = 2;

    private static readonly char[] InvalidSheetNameCharacters =
      { ':', '\\', '/', '?', '*', '[', ']' };

    internal static string SuggestSheetName(string panelName)
    {
      char[] cleaned = (panelName ?? string.Empty)
        .Select(c => char.IsControl(c) ||
                     Array.IndexOf(InvalidSheetNameCharacters, c) >= 0
          ? '-'
          : c)
        .ToArray();
      string name = new string(cleaned).Trim().Trim('\'').Trim();
      if (name.Length > MaximumSheetNameLength)
      {
        name = name.Substring(0, MaximumSheetNameLength).TrimEnd();
      }
      return name;
    }

    internal static bool TryValidateSheetName(
      string name,
      IEnumerable<string> existingNames,
      out string error)
    {
      error = null;
      if (string.IsNullOrWhiteSpace(name))
      {
        error = "Enter a name for the new worksheet.";
      }
      else if (name.Length > MaximumSheetNameLength)
      {
        error = $"Excel sheet names are limited to {MaximumSheetNameLength} characters.";
      }
      else if (name.IndexOfAny(InvalidSheetNameCharacters) >= 0)
      {
        error = "A sheet name cannot contain any of these characters: : \\ / ? * [ ]";
      }
      else if (name.StartsWith("'", StringComparison.Ordinal) ||
               name.EndsWith("'", StringComparison.Ordinal))
      {
        error = "A sheet name cannot begin or end with an apostrophe.";
      }
      else if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
      {
        error = "\"History\" is reserved by Excel; choose another sheet name.";
      }
      else if (existingNames != null &&
               existingNames.Any(existing => string.Equals(
                 existing,
                 name,
                 StringComparison.OrdinalIgnoreCase)))
      {
        error = $"The workbook already has a sheet named \"{name}\".";
      }
      return error == null;
    }

    internal static string DefaultBreakerPhases(string phaseText)
    {
      return int.TryParse(
          (phaseText ?? string.Empty).Trim(),
          NumberStyles.Integer,
          CultureInfo.InvariantCulture,
          out int phases) && phases == 3
        ? "A, B, C"
        : "A, B";
    }

    internal static bool TryParseBreakerPhases(
      string text,
      out string[] phases,
      out string error)
    {
      phases = null;
      error = null;
      List<string> parsed = new List<string>();
      foreach (char c in text ?? string.Empty)
      {
        if (char.IsWhiteSpace(c) || c == ',' || c == ';' || c == '/' || c == '-')
        {
          continue;
        }
        char letter = char.ToUpperInvariant(c);
        if (letter != 'A' && letter != 'B' && letter != 'C')
        {
          error = "Breaker phases must be A, B, or C, for example A, B, C.";
          return false;
        }
        parsed.Add(letter.ToString());
      }

      if (parsed.Count == 0)
      {
        error = "Enter the breaker phase order, for example A, B, C.";
        return false;
      }
      phases = parsed.ToArray();
      return true;
    }

    internal static bool TryParseCircuitCount(
      string text,
      int maximum,
      out int count,
      out string error)
    {
      error = null;
      if (!int.TryParse(
            (text ?? string.Empty).Trim(),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out count) ||
          count < MinimumCircuitCount ||
          count > maximum ||
          count % 2 != 0)
      {
        error = $"Enter an even number of circuit breakers from " +
          $"{MinimumCircuitCount} to {maximum} (the template has {maximum}).";
        return false;
      }
      return true;
    }

    /// <summary>"225" becomes "225A"; anything else is left as typed.</summary>
    internal static string NormalizeAmps(string value)
    {
      string text = (value ?? string.Empty).Trim();
      return Regex.IsMatch(text, @"^\d+$") ? text + "A" : text;
    }

    internal static string ConditionPrefix(string condition)
    {
      switch ((condition ?? string.Empty).Trim().ToUpperInvariant())
      {
        case "NEW": return "(N)";
        case "EXISTING": return "(E)";
        case "RELOCATED": return "(RL)";
        default: return string.Empty;
      }
    }

    internal static string BuildPanelTitle(
      string panelName,
      string condition,
      char quote)
    {
      string prefix = ConditionPrefix(condition);
      string title = $"PANEL {quote}{(panelName ?? string.Empty).Trim()}{quote}";
      return prefix.Length == 0 ? title : prefix + " " + title;
    }
  }

  internal static partial class PanelScheduleWorkbookAllocator
  {
    private const int HeaderScanColumns = 26;

    private sealed class HeaderCell
    {
      internal PanelHeaderField Field;
      internal string Label;
      internal string LabelAddress;
      internal string ValueAddress;
    }

    // Where the ACIES panel schedule keeps each labeled header value. The
    // value of WIRE, PHASE, and TYPE is two columns right of its label, so
    // the position is fixed here and a field is only used when its label is
    // found where it is expected.
    private static readonly HeaderCell[] StandardHeaderCells =
    {
      new HeaderCell { Field = PanelHeaderField.Voltage, Label = "VOLTAGE", LabelAddress = "F2", ValueAddress = "G2" },
      new HeaderCell { Field = PanelHeaderField.BusRating, Label = "BUS RATING", LabelAddress = "F3", ValueAddress = "G3" },
      new HeaderCell { Field = PanelHeaderField.Main, Label = "MAIN REQUIREMENT", LabelAddress = "F4", ValueAddress = "G4" },
      new HeaderCell { Field = PanelHeaderField.Wires, Label = "WIRE", LabelAddress = "I2", ValueAddress = "K2" },
      new HeaderCell { Field = PanelHeaderField.Phase, Label = "PHASE", LabelAddress = "I3", ValueAddress = "K3" },
      new HeaderCell { Field = PanelHeaderField.Type, Label = "TYPE", LabelAddress = "I4", ValueAddress = "K4" },
      new HeaderCell { Field = PanelHeaderField.Mounting, Label = "MOUNTING", LabelAddress = "M2", ValueAddress = "N2" },
      new HeaderCell { Field = PanelHeaderField.AicRating, Label = "AIC RATING", LabelAddress = "M3", ValueAddress = "N3" },
      new HeaderCell { Field = PanelHeaderField.Condition, Label = "CONDITION", LabelAddress = "M4", ValueAddress = "N4" },
    };

    /// <summary>
    /// Lists the panel worksheets in a workbook with their header values and
    /// circuit capacity, opening the file read-only.
    /// </summary>
    internal static PanelTemplateCatalog ReadPanelTemplates(string workbookPath)
    {
      return ExecuteWithWorkbook(workbookPath, true, workbook =>
      {
        PanelTemplateCatalog catalog = new PanelTemplateCatalog();
        Excel.Sheets worksheets = null;
        try
        {
          worksheets = workbook.Worksheets;
          for (int index = 1; index <= worksheets.Count; index++)
          {
            Excel.Worksheet worksheet = null;
            try
            {
              worksheet = worksheets.Item[index] as Excel.Worksheet;
              if (worksheet == null)
              {
                continue;
              }

              catalog.AllSheetNames.Add(worksheet.Name);
              if (!IsPanelSchedule(worksheet))
              {
                continue;
              }

              int capacity = CountCircuitPositions(worksheet);
              if (capacity < PanelSheetRules.MinimumCircuitCount)
              {
                continue;
              }

              catalog.Templates.Add(new PanelTemplateInfo
              {
                SheetName = worksheet.Name,
                CircuitCapacity = capacity,
                Header = ReadHeaderValues(worksheet),
              });
            }
            finally
            {
              ReleaseComObject(worksheet);
            }
          }
        }
        finally
        {
          ReleaseComObject(worksheets);
        }
        return catalog;
      });
    }

    /// <summary>
    /// Copies the template worksheet to the end of the workbook, then applies
    /// the requested header values, circuit count, and breaker phases.
    /// </summary>
    internal static PanelSheetCreationResult CreatePanelWorksheet(
      string workbookPath,
      PanelSheetRequest request)
    {
      if (request == null)
      {
        throw new ArgumentNullException(nameof(request));
      }
      if (!PanelSheetRules.TryValidateSheetName(
        request.NewSheetName,
        null,
        out string nameError))
      {
        throw new InvalidOperationException(nameError);
      }
      if (request.CircuitCount < PanelSheetRules.MinimumCircuitCount ||
          request.CircuitCount % 2 != 0)
      {
        throw new InvalidOperationException(
          "Panel circuit count must be an even number of at least two.");
      }
      if (request.BreakerPhases == null || request.BreakerPhases.Count == 0)
      {
        throw new InvalidOperationException(
          "At least one breaker phase is required.");
      }

      return ExecuteWithWorkbook(workbookPath, false, workbook =>
      {
        Excel.Sheets worksheets = null;
        Excel.Worksheet template = null;
        Excel.Worksheet insertAfter = null;
        Excel.Worksheet created = null;
        try
        {
          worksheets = workbook.Worksheets;
          HashSet<string> existingNames = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
          for (int index = 1; index <= worksheets.Count; index++)
          {
            Excel.Worksheet candidate = null;
            bool keepCandidate = false;
            try
            {
              candidate = worksheets.Item[index] as Excel.Worksheet;
              if (candidate == null)
              {
                continue;
              }
              string candidateName = candidate.Name;
              existingNames.Add(candidateName);
              if (string.Equals(
                candidateName,
                request.NewSheetName,
                StringComparison.OrdinalIgnoreCase))
              {
                throw new InvalidOperationException(
                  $"The workbook already has a sheet named \"{request.NewSheetName}\".");
              }
              if (template == null &&
                  string.Equals(
                    candidateName,
                    request.TemplateSheetName,
                    StringComparison.OrdinalIgnoreCase))
              {
                template = candidate;
                keepCandidate = true;
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

          if (template == null)
          {
            throw new InvalidOperationException(
              $"The template worksheet \"{request.TemplateSheetName}\" was not found.");
          }
          if (!IsPanelSchedule(template))
          {
            throw new InvalidOperationException(
              $"\"{request.TemplateSheetName}\" is not a panel schedule worksheet.");
          }

          // Excel puts the copy after the last visible sheet, so a workbook
          // that ends with hidden sheets does not end with the copy. Find the
          // copy by its name rather than by its position.
          insertAfter = FindLastVisibleWorksheet(worksheets);
          template.Copy(Type.Missing, insertAfter ?? template);
          created = FindWorksheetNotIn(worksheets, existingNames);
          if (created == null)
          {
            throw new InvalidOperationException(
              "Excel did not add exactly one new worksheet when copying the template.");
          }

          created.Name = request.NewSheetName;
          created.Visible = Excel.XlSheetVisibility.xlSheetVisible;

          PanelSheetCreationResult result = new PanelSheetCreationResult
          {
            SheetName = created.Name,
            CircuitCount = request.CircuitCount,
          };
          ConfigureNewPanelWorksheet(created, request, result);
          try
          {
            created.Activate();
          }
          catch
          {
            // The workbook still saves correctly if Excel cannot activate it.
          }
          return result;
        }
        finally
        {
          ReleaseComObject(created);
          ReleaseComObject(insertAfter);
          ReleaseComObject(template);
          ReleaseComObject(worksheets);
        }
      });
    }

    private static Excel.Worksheet FindLastVisibleWorksheet(Excel.Sheets worksheets)
    {
      for (int index = worksheets.Count; index >= 1; index--)
      {
        Excel.Worksheet candidate = worksheets.Item[index] as Excel.Worksheet;
        if (candidate != null &&
            candidate.Visible == Excel.XlSheetVisibility.xlSheetVisible)
        {
          return candidate;
        }
        ReleaseComObject(candidate);
      }
      return null;
    }

    /// <summary>
    /// Returns the single worksheet whose name is not in the given set, or
    /// null when there is not exactly one.
    /// </summary>
    private static Excel.Worksheet FindWorksheetNotIn(
      Excel.Sheets worksheets,
      HashSet<string> knownNames)
    {
      Excel.Worksheet found = null;
      for (int index = 1; index <= worksheets.Count; index++)
      {
        Excel.Worksheet candidate = worksheets.Item[index] as Excel.Worksheet;
        if (candidate == null || knownNames.Contains(candidate.Name))
        {
          ReleaseComObject(candidate);
          continue;
        }
        if (found != null)
        {
          ReleaseComObject(candidate);
          ReleaseComObject(found);
          return null;
        }
        found = candidate;
      }
      return found;
    }

    private static void ConfigureNewPanelWorksheet(
      Excel.Worksheet worksheet,
      PanelSheetRequest request,
      PanelSheetCreationResult result)
    {
      if (!TryFindCircuitStartRow(worksheet, out int startRow))
      {
        throw new InvalidOperationException(
          "The template worksheet does not contain numbered circuits.");
      }

      int templateCapacity = CountCircuitPositions(worksheet);
      if (request.CircuitCount > templateCapacity)
      {
        throw new InvalidOperationException(
          $"The template \"{request.TemplateSheetName}\" only has " +
          $"{templateCapacity} circuit positions, not {request.CircuitCount}.");
      }

      string rightCircuitColumn = DetectRightCircuitColumn(worksheet, startRow);
      int phaseColumn = FindPhaseColumn(worksheet);
      int rowCount = templateCapacity / 2;

      ClearCircuitEntries(
        worksheet,
        startRow,
        rowCount,
        ColumnIndex(rightCircuitColumn),
        phaseColumn);

      for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
      {
        WriteCell(
          worksheet,
          $"{ColumnLetters(phaseColumn)}{startRow + rowIndex}",
          request.BreakerPhases[rowIndex % request.BreakerPhases.Count]);
      }

      // Rows past the requested count are hidden, which is also how SPS later
      // detects the panel's circuit capacity.
      List<PanelCircuitSlot> slots = BuildCircuitSlots(worksheet);
      ApplyCapacityAndSpareRules(worksheet, slots, request.CircuitCount, 0);

      WritePanelTitle(worksheet, request);
      Dictionary<PanelHeaderField, string> addresses =
        FindHeaderValueAddresses(worksheet);
      foreach (PanelHeaderField field in Enum.GetValues(typeof(PanelHeaderField)))
      {
        if (!addresses.TryGetValue(field, out string address))
        {
          result.MissingFields.Add(field);
          continue;
        }
        WriteCell(worksheet, address, ToHeaderCellValue(field, request.Header));
      }
    }

    private static object ToHeaderCellValue(
      PanelHeaderField field,
      PanelHeaderValues header)
    {
      string text = (header.Get(field) ?? string.Empty).Trim();
      if (field == PanelHeaderField.BusRating)
      {
        text = PanelSheetRules.NormalizeAmps(text);
      }
      if (text.Length == 0)
      {
        return null;
      }

      // The template stores wire and phase counts as numbers.
      if ((field == PanelHeaderField.Wires || field == PanelHeaderField.Phase) &&
          int.TryParse(
            text,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out int number))
      {
        return (double)number;
      }

      // Keep Excel from reading typed text as a formula or a date.
      if (text[0] == '=' ||
          (DateTime.TryParse(
             text,
             CultureInfo.CurrentCulture,
             DateTimeStyles.None,
             out _) &&
           !double.TryParse(
             text,
             NumberStyles.Any,
             CultureInfo.CurrentCulture,
             out _)))
      {
        return "'" + text;
      }
      return text;
    }

    private static void WritePanelTitle(
      Excel.Worksheet worksheet,
      PanelSheetRequest request)
    {
      string titleAddress = "A3";
      char quote = '\'';
      for (int row = 1; row <= 4; row++)
      {
        string text = ReadText(worksheet, $"A{row}");
        Match titleMatch = Regex.Match(
          text,
          @"\b(?:PANEL|PNL)\s*([\""'])",
          RegexOptions.IgnoreCase);
        if (titleMatch.Success)
        {
          titleAddress = $"A{row}";
          quote = titleMatch.Groups[1].Value[0];
          break;
        }
        if (Regex.IsMatch(text, @"\b(?:PANEL|PNL)\b", RegexOptions.IgnoreCase))
        {
          titleAddress = $"A{row}";
          break;
        }
      }

      WriteCell(
        worksheet,
        titleAddress,
        PanelSheetRules.BuildPanelTitle(
          request.PanelName,
          request.Header.Condition,
          quote));
    }

    private static PanelHeaderValues ReadHeaderValues(Excel.Worksheet worksheet)
    {
      PanelHeaderValues header = new PanelHeaderValues();
      foreach (KeyValuePair<PanelHeaderField, string> entry in
        FindHeaderValueAddresses(worksheet))
      {
        header.Set(
          entry.Key,
          Regex.Replace(ReadText(worksheet, entry.Value), @"\s+", " ").Trim());
      }
      return header;
    }

    /// <summary>
    /// The value cell of each header field whose label (for example
    /// "VOLTAGE:") is where the standard layout puts it. A field whose label
    /// is missing is left out so nothing is written to a guessed cell.
    /// </summary>
    private static Dictionary<PanelHeaderField, string> FindHeaderValueAddresses(
      Excel.Worksheet worksheet)
    {
      Dictionary<PanelHeaderField, string> addresses =
        new Dictionary<PanelHeaderField, string>();
      foreach (HeaderCell cell in StandardHeaderCells)
      {
        string label = Regex.Replace(
          ReadText(worksheet, cell.LabelAddress),
          @"\s+",
          " ").Trim().TrimEnd(':').Trim();
        if (string.Equals(
          label,
          cell.Label,
          StringComparison.OrdinalIgnoreCase))
        {
          addresses[cell.Field] = cell.ValueAddress;
        }
      }
      return addresses;
    }

    private static int FindTableHeaderRow(Excel.Worksheet worksheet)
    {
      if (string.Equals(
        ReadText(worksheet, "A5"),
        "CKT#",
        StringComparison.OrdinalIgnoreCase))
      {
        return 5;
      }
      return string.Equals(
        ReadText(worksheet, "A6"),
        "CKT#",
        StringComparison.OrdinalIgnoreCase)
        ? 6
        : 0;
    }

    private static int FindPhaseColumn(Excel.Worksheet worksheet)
    {
      int headerRow = FindTableHeaderRow(worksheet);
      if (headerRow > 0)
      {
        Excel.Range headers = null;
        try
        {
          headers = worksheet.Range[
            $"A{headerRow}:{ColumnLetters(HeaderScanColumns)}{headerRow}"];
          object[,] values = headers.Value2 as object[,];
          if (values != null)
          {
            for (int column = 1; column <= values.GetLength(1); column++)
            {
              if (string.Equals(
                (Convert.ToString(
                  values[1, column],
                  CultureInfo.InvariantCulture) ?? string.Empty).Trim(),
                "PHASE",
                StringComparison.OrdinalIgnoreCase))
              {
                return column;
              }
            }
          }
        }
        finally
        {
          ReleaseComObject(headers);
        }
      }

      // Every ACIES schedule keeps the phase letters in column J.
      return 10;
    }

    /// <summary>
    /// Counts circuit positions in a panel worksheet whether or not their rows
    /// are hidden, rounded up to an even number.
    /// </summary>
    private static int CountCircuitPositions(Excel.Worksheet worksheet)
    {
      if (!TryFindCircuitStartRow(worksheet, out int startRow))
      {
        return 0;
      }

      string rightColumn = DetectRightCircuitColumn(worksheet, startRow);
      const int maximumRows = 42;
      int endRow = startRow + maximumRows - 1;
      Excel.Range leftRange = null;
      Excel.Range rightRange = null;
      try
      {
        leftRange = worksheet.Range[$"A{startRow}:A{endRow}"];
        rightRange = worksheet.Range[$"{rightColumn}{startRow}:{rightColumn}{endRow}"];
        object[,] left = leftRange.Value2 as object[,];
        object[,] right = rightRange.Value2 as object[,];
        if (left == null || right == null)
        {
          return 0;
        }

        int highest = 0;
        for (int index = 1; index <= maximumRows; index++)
        {
          int leftCircuit = ToCircuitNumber(left[index, 1]);
          int rightCircuit = ToCircuitNumber(right[index, 1]);
          bool hasLeft = IsValidCircuitNumber(leftCircuit);
          bool hasRight = IsValidCircuitNumber(rightCircuit);
          if (!hasLeft && !hasRight)
          {
            break;
          }
          highest = Math.Max(
            highest,
            Math.Max(
              hasLeft ? leftCircuit : rightCircuit - 1,
              hasRight ? rightCircuit : leftCircuit + 1));
        }
        return highest % 2 == 0 ? highest : highest + 1;
      }
      finally
      {
        ReleaseComObject(rightRange);
        ReleaseComObject(leftRange);
      }
    }

    private static int ToCircuitNumber(object value)
    {
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
        return 0;
      }
    }

    /// <summary>
    /// Empties the load columns of every circuit row, keeping the circuit
    /// numbers on both sides and the phase column.
    /// </summary>
    private static void ClearCircuitEntries(
      Excel.Worksheet worksheet,
      int startRow,
      int rowCount,
      int rightCircuitColumn,
      int phaseColumn)
    {
      const int firstColumn = 2;
      int lastColumn = rightCircuitColumn - 1;
      if (rowCount < 1 || lastColumn < firstColumn)
      {
        return;
      }

      Excel.Range block = null;
      try
      {
        block = worksheet.Range[
          $"{ColumnLetters(firstColumn)}{startRow}:" +
          $"{ColumnLetters(lastColumn)}{startRow + rowCount - 1}"];
        object[,] formulas = block.Formula as object[,];
        if (formulas == null)
        {
          return;
        }

        for (int row = 1; row <= formulas.GetLength(0); row++)
        {
          for (int offset = 1; offset <= formulas.GetLength(1); offset++)
          {
            int column = firstColumn + offset - 1;
            if (column == phaseColumn ||
                string.IsNullOrEmpty(
                  Convert.ToString(
                    formulas[row, offset],
                    CultureInfo.InvariantCulture)))
            {
              continue;
            }
            ClearCell(
              worksheet,
              $"{ColumnLetters(column)}{startRow + row - 1}");
          }
        }
      }
      finally
      {
        ReleaseComObject(block);
      }
    }

    private static void ClearCell(Excel.Worksheet worksheet, string address)
    {
      Excel.Range range = null;
      Excel.Range mergeArea = null;
      try
      {
        range = worksheet.Range[address];
        if (Convert.ToBoolean(range.MergeCells))
        {
          mergeArea = range.MergeArea;
          string topLeft = ((Excel.Range)mergeArea.Cells[1, 1])
            .Address[false, false];
          if (!string.Equals(topLeft, address, StringComparison.OrdinalIgnoreCase))
          {
            // The merged area is cleared once, from its top-left cell.
            return;
          }
          mergeArea.ClearContents();
          return;
        }
        range.ClearContents();
      }
      finally
      {
        ReleaseComObject(mergeArea);
        ReleaseComObject(range);
      }
    }

    private static int ColumnIndex(string letters)
    {
      int index = 0;
      foreach (char c in (letters ?? string.Empty).ToUpperInvariant())
      {
        index = index * 26 + (c - 'A' + 1);
      }
      return index;
    }

    private static string ColumnLetters(int index)
    {
      string letters = string.Empty;
      while (index > 0)
      {
        int remainder = (index - 1) % 26;
        letters = (char)('A' + remainder) + letters;
        index = (index - 1) / 26;
      }
      return letters;
    }
  }

  public partial class GeneralCommands
  {
    /// <summary>
    /// Offers to create the missing panel worksheet from a template sheet.
    /// Returns true when a worksheet was created and saved.
    /// </summary>
    private static bool TryCreatePanelWorksheet(
      Editor editor,
      string workbookPath,
      string panelName,
      PanelWorksheetNotFoundException notFound)
    {
      string workbookName = Path.GetFileName(workbookPath);
      editor.WriteMessage(
        $"\nPanel {panelName} is not in \"{workbookName}\". " +
        $"Existing panel worksheets: " +
        $"{(notFound.AvailableWorksheets.Count == 0 ? "none" : string.Join(", ", notFound.AvailableWorksheets))}.");

      // Saving the new worksheet needs the file to be writable, so find out
      // before the user fills in the dialog.
      if (!TryVerifyPanelScheduleWorkbookClosed(workbookPath, out _))
      {
        editor.WriteMessage(
          "\nThe workbook is open in Excel or is read-only, so a worksheet " +
          "cannot be added to it. Close it in Excel, then run SPS again.");
        return false;
      }

      PanelTemplateCatalog catalog;
      try
      {
        catalog = PanelScheduleWorkbookAllocator.ReadPanelTemplates(
          workbookPath);
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage(
          $"\nUnable to read the panel worksheets in the workbook: {ex.Message}");
        return false;
      }

      if (catalog.Templates.Count == 0)
      {
        editor.WriteMessage(
          "\nThe workbook has no panel worksheet to copy as a template. " +
          "Choose a workbook that contains one (for example a TEMPLATE sheet).");
        return false;
      }

      PanelSheetCreateWindow window = new PanelSheetCreateWindow(
        panelName,
        workbookName,
        catalog);
      bool? accepted =
        Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(window);
      if (accepted != true || window.Request == null)
      {
        return false;
      }

      PanelSheetCreationResult result =
        PanelScheduleWorkbookAllocator.CreatePanelWorksheet(
          workbookPath,
          window.Request);
      editor.WriteMessage(
        $"\nCreated worksheet \"{result.SheetName}\" from template " +
        $"\"{window.Request.TemplateSheetName}\" with {result.CircuitCount} circuits.");
      if (result.MissingFields.Count > 0)
      {
        editor.WriteMessage(
          "\nThe template has no cell for: " +
          string.Join(", ", result.MissingFields) +
          ". Those were left as the template had them.");
      }
      return true;
    }
  }
}
