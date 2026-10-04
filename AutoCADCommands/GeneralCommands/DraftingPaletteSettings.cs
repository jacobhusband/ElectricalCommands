using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ElectricalCommands
{
  /// <summary>
  /// Per-user preferences for the drafting palette, kept in
  /// %APPDATA%\ElectricalCommands so they apply to every drawing. Nothing here
  /// touches AutoCAD, so the startup hook can read it before any UI exists.
  /// </summary>
  internal static class DraftingPaletteSettings
  {
    private const string ShowAtStartupKey = "showAtStartup";

    internal static string DefaultFilePath =>
      Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "ElectricalCommands",
        "drafting-palette.json");

    internal static bool ReadShowAtStartup()
    {
      return ReadShowAtStartup(DefaultFilePath);
    }

    // The palette opens at startup unless the user turned that off, so a
    // missing, unreadable, or malformed setting means "on".
    internal static bool ReadShowAtStartup(string filePath)
    {
      try
      {
        JToken value = Load(filePath)?[ShowAtStartupKey];
        return value == null || value.Type != JTokenType.Boolean || value.Value<bool>();
      }
      catch
      {
        return true;
      }
    }

    internal static bool TryWriteShowAtStartup(bool value, out string error)
    {
      return TryWriteShowAtStartup(DefaultFilePath, value, out error);
    }

    internal static bool TryWriteShowAtStartup(string filePath, bool value, out string error)
    {
      error = string.Empty;
      try
      {
        // Keep any other keys already in the file.
        JObject root = Load(filePath) ?? new JObject();
        root[ShowAtStartupKey] = value;
        Directory.CreateDirectory(Path.GetDirectoryName(filePath));
        File.WriteAllText(filePath, root.ToString(Formatting.Indented));
        return true;
      }
      catch (Exception ex)
      {
        error = ex.Message;
        return false;
      }
    }

    private static JObject Load(string filePath)
    {
      if (!File.Exists(filePath))
      {
        return null;
      }

      try
      {
        return JObject.Parse(File.ReadAllText(filePath));
      }
      catch (JsonException)
      {
        return null;
      }
    }
  }
}
