using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace ElectricalCommands
{
  /// <summary>How a note tab's table is numbered.</summary>
  public enum NoteTableKind
  {
    /// <summary>Hexagon KEYED_NOTE symbols, as KNT draws them.</summary>
    Keyed,

    /// <summary>Plain "1." numbers, as GNT draws them.</summary>
    General,
  }

  internal sealed class NoteTableProject
  {
    public int Version { get; set; } = 1;
    public List<NoteTableTabData> Tabs { get; set; } = new List<NoteTableTabData>();
  }

  internal sealed class NoteTableTabData
  {
    public string Name { get; set; } = string.Empty;

    [JsonConverter(typeof(StringEnumConverter))]
    public NoteTableKind Kind { get; set; } = NoteTableKind.Keyed;

    public int FirstNumber { get; set; } = 1;
    public List<string> Notes { get; set; } = new List<string>();
  }

  /// <summary>
  /// Reads and writes the NOTETABLE tabs, kept in one JSON file beside the
  /// drawings so every sheet in the project folder shares them.
  /// </summary>
  internal static class NoteTableStore
  {
    internal const string FileName = "NoteTables.json";
    internal const string GeneralNotesTabName = "General Notes";
    internal const string KeyNotesTabName = "Key Notes";

    // Checked in order, so "E-DEMO POWER" is a demolition plan. Short words
    // must match a whole word of the file name; longer ones may match inside
    // one ("ELIGHTING", "SITEPLAN").
    private static readonly (string TabName, string[] Words)[] PlanTypes =
    {
      ("Demolition Plan", new[] { "demolition", "demo" }),
      ("Site Plan", new[] { "site" }),
      ("Roof Plan", new[] { "roof" }),
      ("Lighting Plan", new[] { "lighting", "light", "ltg", "lt" }),
      ("Power Plan", new[] { "power", "pwr", "pw" }),
      ("Systems Plan", new[] { "systems", "signal", "sys" }),
      ("Fire Alarm Plan", new[] { "firealarm", "fire", "fa" }),
      ("Single Line Diagram", new[] { "singleline", "oneline", "sld" }),
      ("Enlarged Plan", new[] { "enlarged" }),
    };

    private const int WholeWordLength = 3;

    /// <summary>
    /// The project file for a saved drawing, or null when the drawing has
    /// never been saved and so has no folder.
    /// </summary>
    internal static string GetProjectFilePath(string drawingPath)
    {
      if (string.IsNullOrWhiteSpace(drawingPath) ||
          !Path.IsPathRooted(drawingPath) ||
          !File.Exists(drawingPath))
      {
        return null;
      }
      string folder = Path.GetDirectoryName(drawingPath);
      return string.IsNullOrEmpty(folder) ? null : Path.Combine(folder, FileName);
    }

    /// <summary>
    /// Loads the project file; a missing file is an empty project. A file that
    /// cannot be read throws rather than being treated as empty, so a later
    /// save cannot wipe out notes that are still in it.
    /// </summary>
    internal static NoteTableProject Load(string path)
    {
      if (string.IsNullOrEmpty(path) || !File.Exists(path))
      {
        return new NoteTableProject();
      }

      try
      {
        NoteTableProject project = JsonConvert.DeserializeObject<NoteTableProject>(
          File.ReadAllText(path));
        return Normalize(project ?? new NoteTableProject());
      }
      catch (JsonException ex)
      {
        throw new InvalidDataException(
          $"{path} is not a valid note table file: {ex.Message}",
          ex);
      }
    }

    internal static void Save(string path, NoteTableProject project)
    {
      string json = JsonConvert.SerializeObject(Normalize(project), Formatting.Indented);
      // Write beside the target first so a failed write never leaves a
      // half-written project file.
      string temporaryPath = path + ".tmp";
      File.WriteAllText(temporaryPath, json);
      if (File.Exists(path))
      {
        File.Replace(temporaryPath, path, null);
      }
      else
      {
        File.Move(temporaryPath, path);
      }
    }

    internal static DateTime GetLastWriteTimeUtc(string path)
    {
      return !string.IsNullOrEmpty(path) && File.Exists(path)
        ? File.GetLastWriteTimeUtc(path)
        : DateTime.MinValue;
    }

    /// <summary>
    /// The plan tab name for a drawing ("Power Plan" for E-POWER.dwg), or
    /// null when the file name names no plan type.
    /// </summary>
    internal static string DetectPlanTabName(string drawingPath)
    {
      string name = Path.GetFileNameWithoutExtension(drawingPath ?? string.Empty)
        .ToLowerInvariant();
      if (name.Length == 0)
      {
        return null;
      }

      string[] words = name
        .Split(name.Where(ch => !char.IsLetter(ch)).Distinct().ToArray(),
          StringSplitOptions.RemoveEmptyEntries);
      string joined = string.Concat(words);
      foreach ((string tabName, string[] planWords) in PlanTypes)
      {
        foreach (string word in planWords)
        {
          bool found = word.Length <= WholeWordLength
            ? words.Contains(word)
            : words.Any(part => part.Contains(word)) || joined.Contains(word);
          if (found)
          {
            return tabName;
          }
        }
      }
      return null;
    }

    /// <summary>
    /// Adds the tabs a drawing should start with when they are missing: General
    /// Notes, and a keyed-note tab for the plan the drawing name describes (or
    /// a plain Key Notes tab when it names none and no keyed tab exists yet).
    /// </summary>
    internal static void AddDefaultTabs(NoteTableProject project, string planTabName)
    {
      if (project.Tabs.Count == 0)
      {
        project.Tabs.Add(new NoteTableTabData
        {
          Name = GeneralNotesTabName,
          Kind = NoteTableKind.General,
        });
      }

      if (!string.IsNullOrEmpty(planTabName))
      {
        if (FindTab(project, planTabName) == null)
        {
          project.Tabs.Add(new NoteTableTabData
          {
            Name = planTabName,
            Kind = NoteTableKind.Keyed,
          });
        }
      }
      else if (!project.Tabs.Any(tab => tab.Kind == NoteTableKind.Keyed))
      {
        project.Tabs.Add(new NoteTableTabData
        {
          Name = KeyNotesTabName,
          Kind = NoteTableKind.Keyed,
        });
      }
    }

    internal static NoteTableTabData FindTab(NoteTableProject project, string name)
    {
      return project.Tabs.FirstOrDefault(tab => string.Equals(
        tab.Name,
        name,
        StringComparison.OrdinalIgnoreCase));
    }

    private static NoteTableProject Normalize(NoteTableProject project)
    {
      project.Tabs = (project.Tabs ?? new List<NoteTableTabData>())
        .Where(tab => tab != null)
        .ToList();
      foreach (NoteTableTabData tab in project.Tabs)
      {
        tab.Name = (tab.Name ?? string.Empty).Trim();
        tab.FirstNumber = Math.Max(1, tab.FirstNumber);
        tab.Notes = (tab.Notes ?? new List<string>())
          .Select(NormalizeNoteText)
          .Where(note => note.Length > 0)
          .ToList();
      }
      return project;
    }

    internal static string NormalizeNoteText(string text)
    {
      return string.Join(
          "\n",
          (text ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Split('\n')
            .Select(line => line.TrimEnd()))
        .Trim();
    }
  }
}
