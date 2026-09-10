using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  internal static class AddNoteStore
  {
    private const string NotesFileName = "notes.json";
    private static readonly Regex NoteSplitRegex =
      new Regex(@"\r?\n\s*\r?\n", RegexOptions.Compiled);

    internal static AddNoteLoadResult Load()
    {
      string notesPath = Path.Combine(ProjectChecklistStore.ResolveAppDataFolder(), NotesFileName);
      if (!File.Exists(notesPath))
      {
        return AddNoteLoadResult.Empty(
          notesPath,
          $"Notes file was not found at {notesPath}."
        );
      }

      AddNoteFile payload;
      try
      {
        string json = File.ReadAllText(notesPath, Encoding.UTF8);
        payload = JsonConvert.DeserializeObject<AddNoteFile>(json) ?? new AddNoteFile();
      }
      catch (System.Exception ex)
      {
        return AddNoteLoadResult.Empty(
          notesPath,
          $"Notes file could not be read at {notesPath}: {ex.Message}"
        );
      }

      List<AddNoteTab> tabs = NormalizeTabs(payload);
      if (tabs.Count == 0 || tabs.All(tab => tab.Notes.Count == 0))
      {
        return AddNoteLoadResult.Empty(notesPath, $"No notes were found in {notesPath}.");
      }

      return new AddNoteLoadResult
      {
        NotesPath = notesPath,
        Tabs = tabs,
        StatusMessage = $"Loaded notes from {notesPath}.",
      };
    }

    private static List<AddNoteTab> NormalizeTabs(AddNoteFile payload)
    {
      var tabNames = new List<string>();
      var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

      foreach (string rawTab in payload?.Tabs ?? new List<string>())
      {
        AddTabName(tabNames, seen, rawTab);
      }

      foreach (string rawTab in (payload?.Keyed ?? new Dictionary<string, string>()).Keys)
      {
        AddTabName(tabNames, seen, rawTab);
      }

      foreach (string rawTab in (payload?.General ?? new Dictionary<string, string>()).Keys)
      {
        AddTabName(tabNames, seen, rawTab);
      }

      if (tabNames.Count == 0)
      {
        tabNames.Add("General");
      }

      return tabNames
        .Select(tabName => new AddNoteTab
        {
          Name = tabName,
          Notes = SplitNotes(ResolveTabContent(payload, tabName)),
        })
        .Where(tab => tab.Notes.Count > 0)
        .ToList();
    }

    private static void AddTabName(
      List<string> tabNames,
      HashSet<string> seen,
      string rawTab
    )
    {
      string tabName = NormalizePlainText(rawTab);
      if (string.IsNullOrWhiteSpace(tabName) || seen.Contains(tabName))
      {
        return;
      }

      seen.Add(tabName);
      tabNames.Add(tabName);
    }

    private static string ResolveTabContent(AddNoteFile payload, string tabName)
    {
      string keyedContent = GetTrimmedValue(payload?.Keyed, tabName);
      string generalContent = GetTrimmedValue(payload?.General, tabName);

      if (!string.IsNullOrWhiteSpace(keyedContent) && !string.IsNullOrWhiteSpace(generalContent))
      {
        return $"{keyedContent}\n\n--- General Notes ---\n\n{generalContent}";
      }

      return !string.IsNullOrWhiteSpace(keyedContent) ? keyedContent : generalContent;
    }

    private static string GetTrimmedValue(Dictionary<string, string> values, string tabName)
    {
      if (values == null || string.IsNullOrWhiteSpace(tabName))
      {
        return string.Empty;
      }

      foreach (KeyValuePair<string, string> pair in values)
      {
        if (string.Equals(pair.Key, tabName, StringComparison.OrdinalIgnoreCase))
        {
          return (pair.Value ?? string.Empty).Trim();
        }
      }

      return string.Empty;
    }

    private static List<AddNoteItem> SplitNotes(string content)
    {
      return NoteSplitRegex
        .Split(content ?? string.Empty)
        .Select(NormalizeNoteText)
        .Where(note => !string.IsNullOrWhiteSpace(note))
        .Select(note => new AddNoteItem { Text = note })
        .ToList();
    }

    private static string NormalizeNoteText(string value)
    {
      return (value ?? string.Empty)
        .Replace("\r\n", "\n")
        .Replace("\r", "\n")
        .Trim();
    }

    private static string NormalizePlainText(string value)
    {
      return (value ?? string.Empty).Trim();
    }
  }

  internal sealed class AddNoteFile
  {
    [JsonProperty("tabs")]
    public List<string> Tabs { get; set; } = new List<string>();

    [JsonProperty("keyed")]
    public Dictionary<string, string> Keyed { get; set; } =
      new Dictionary<string, string>();

    [JsonProperty("general")]
    public Dictionary<string, string> General { get; set; } =
      new Dictionary<string, string>();
  }

  internal sealed class AddNoteLoadResult
  {
    public string NotesPath { get; set; }
    public List<AddNoteTab> Tabs { get; set; } = new List<AddNoteTab>();
    public string StatusMessage { get; set; }

    internal static AddNoteLoadResult Empty(string notesPath, string statusMessage)
    {
      return new AddNoteLoadResult
      {
        NotesPath = notesPath,
        StatusMessage = statusMessage,
      };
    }
  }

  internal sealed class AddNoteTab
  {
    public string Name { get; set; }
    public List<AddNoteItem> Notes { get; set; } = new List<AddNoteItem>();
  }

  internal sealed class AddNoteItem
  {
    public string Text { get; set; }

    public string Preview
    {
      get
      {
        string compact = Regex.Replace(Text ?? string.Empty, @"\s+", " ").Trim();
        return compact.Length <= 220 ? compact : compact.Substring(0, 217) + "...";
      }
    }
  }
}
