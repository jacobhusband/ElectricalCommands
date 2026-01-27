using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ElectricalCommands
{
  /// <summary>
  /// Represents a single signature entry in the signature manager.
  /// </summary>
  public class SignatureEntry
  {
    /// <summary>
    /// Unique identifier for the signature.
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// User-friendly display name (e.g., "Wilson Lee").
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Full path to the signature image file.
    /// </summary>
    public string FilePath { get; set; }

    /// <summary>
    /// True if this is the bundled default signature (cannot be deleted).
    /// </summary>
    public bool IsDefault { get; set; }

    /// <summary>
    /// When this signature was added.
    /// </summary>
    public DateTime DateAdded { get; set; }
  }

  /// <summary>
  /// Settings for T24 signature management.
  /// </summary>
  public class T24SignatureSettings
  {
    /// <summary>
    /// ID of the currently selected/active signature.
    /// </summary>
    public string SelectedSignatureId { get; set; }

    /// <summary>
    /// List of all available signatures.
    /// </summary>
    public List<SignatureEntry> Signatures { get; set; } = new List<SignatureEntry>();
  }

  /// <summary>
  /// Static helper class for managing T24 signatures.
  /// </summary>
  public static class SignatureManager
  {
    private const string DEFAULT_SIGNATURE_NAME = "Wilson Lee (Default)";
    private const string DEFAULT_SIGNATURE_FILENAME = "WL_Sig_Blue_Small.gif";
    private const string SETTINGS_FILENAME = "T24SignatureSettings.json";
    private const string SIGNATURES_FOLDER = "signatures";

    /// <summary>
    /// Gets the path to the signatures storage folder.
    /// </summary>
    public static string GetSignaturesFolderPath()
    {
      string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
      string folderPath = Path.Combine(appDataPath, "ElectricalCommands", SIGNATURES_FOLDER);
      if (!Directory.Exists(folderPath))
      {
        Directory.CreateDirectory(folderPath);
      }
      return folderPath;
    }

    /// <summary>
    /// Gets the path to the settings JSON file.
    /// </summary>
    public static string GetSettingsFilePath()
    {
      string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
      string folderPath = Path.Combine(appDataPath, "ElectricalCommands");
      if (!Directory.Exists(folderPath))
      {
        Directory.CreateDirectory(folderPath);
      }
      return Path.Combine(folderPath, SETTINGS_FILENAME);
    }

    /// <summary>
    /// Loads signature settings, initializing defaults if needed.
    /// </summary>
    public static T24SignatureSettings LoadSettings()
    {
      string settingsPath = GetSettingsFilePath();

      T24SignatureSettings settings = null;

      // Try to load existing settings
      if (File.Exists(settingsPath))
      {
        try
        {
          string json = File.ReadAllText(settingsPath);
          settings = JsonConvert.DeserializeObject<T24SignatureSettings>(json);
        }
        catch
        {
          settings = null;
        }
      }

      // Initialize if needed
      if (settings == null || settings.Signatures == null || settings.Signatures.Count == 0)
      {
        settings = InitializeDefaultSettings();
      }

      // Validate that selected signature exists
      if (string.IsNullOrEmpty(settings.SelectedSignatureId) ||
          !settings.Signatures.Any(s => s.Id == settings.SelectedSignatureId))
      {
        // Select the default signature
        var defaultSig = settings.Signatures.FirstOrDefault(s => s.IsDefault);
        if (defaultSig != null)
        {
          settings.SelectedSignatureId = defaultSig.Id;
        }
        else if (settings.Signatures.Count > 0)
        {
          settings.SelectedSignatureId = settings.Signatures[0].Id;
        }
      }

      return settings;
    }

    /// <summary>
    /// Saves signature settings to disk.
    /// </summary>
    public static void SaveSettings(T24SignatureSettings settings)
    {
      try
      {
        string settingsPath = GetSettingsFilePath();
        string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
        File.WriteAllText(settingsPath, json);
      }
      catch
      {
        // Silently fail - settings will be reloaded next time
      }
    }

    /// <summary>
    /// Gets the currently selected signature entry.
    /// </summary>
    public static SignatureEntry GetSelectedSignature(T24SignatureSettings settings)
    {
      if (settings == null || string.IsNullOrEmpty(settings.SelectedSignatureId))
        return null;

      return settings.Signatures.FirstOrDefault(s => s.Id == settings.SelectedSignatureId);
    }

    /// <summary>
    /// Adds a new signature from a source file.
    /// </summary>
    public static SignatureEntry AddSignature(T24SignatureSettings settings, string sourceFilePath, string name)
    {
      if (!File.Exists(sourceFilePath))
        return null;

      // Generate unique ID and filename
      string id = Guid.NewGuid().ToString();
      string extension = Path.GetExtension(sourceFilePath);
      string destFileName = id + extension;
      string destPath = Path.Combine(GetSignaturesFolderPath(), destFileName);

      // Copy file to signatures folder
      try
      {
        File.Copy(sourceFilePath, destPath, true);
      }
      catch
      {
        return null;
      }

      // Create entry
      var entry = new SignatureEntry
      {
        Id = id,
        Name = name,
        FilePath = destPath,
        IsDefault = false,
        DateAdded = DateTime.Now
      };

      settings.Signatures.Add(entry);
      return entry;
    }

    /// <summary>
    /// Removes a signature (cannot remove default).
    /// </summary>
    public static bool RemoveSignature(T24SignatureSettings settings, string signatureId)
    {
      var entry = settings.Signatures.FirstOrDefault(s => s.Id == signatureId);
      if (entry == null || entry.IsDefault)
        return false;

      // Delete the file if it exists
      try
      {
        if (File.Exists(entry.FilePath))
        {
          File.Delete(entry.FilePath);
        }
      }
      catch
      {
        // Continue even if file delete fails
      }

      settings.Signatures.Remove(entry);

      // If this was the selected signature, switch to default
      if (settings.SelectedSignatureId == signatureId)
      {
        var defaultSig = settings.Signatures.FirstOrDefault(s => s.IsDefault);
        settings.SelectedSignatureId = defaultSig?.Id ?? settings.Signatures.FirstOrDefault()?.Id;
      }

      return true;
    }

    /// <summary>
    /// Renames a signature.
    /// </summary>
    public static bool RenameSignature(T24SignatureSettings settings, string signatureId, string newName)
    {
      var entry = settings.Signatures.FirstOrDefault(s => s.Id == signatureId);
      if (entry == null || string.IsNullOrWhiteSpace(newName))
        return false;

      entry.Name = newName.Trim();
      return true;
    }

    /// <summary>
    /// Initializes default settings with the bundled signature.
    /// </summary>
    private static T24SignatureSettings InitializeDefaultSettings()
    {
      var settings = new T24SignatureSettings
      {
        Signatures = new List<SignatureEntry>()
      };

      // Extract bundled default signature
      string defaultPath = ExtractDefaultSignature();

      if (!string.IsNullOrEmpty(defaultPath))
      {
        var defaultEntry = new SignatureEntry
        {
          Id = Guid.NewGuid().ToString(),
          Name = DEFAULT_SIGNATURE_NAME,
          FilePath = defaultPath,
          IsDefault = true,
          DateAdded = DateTime.Now
        };

        settings.Signatures.Add(defaultEntry);
        settings.SelectedSignatureId = defaultEntry.Id;
      }

      SaveSettings(settings);
      return settings;
    }

    /// <summary>
    /// Extracts the bundled default signature to the signatures folder.
    /// </summary>
    private static string ExtractDefaultSignature()
    {
      string destPath = Path.Combine(GetSignaturesFolderPath(), DEFAULT_SIGNATURE_FILENAME);

      // If already exists, return it
      if (File.Exists(destPath))
        return destPath;

      // Try to extract from embedded resource
      try
      {
        var assembly = Assembly.GetExecutingAssembly();

        // Try different possible resource names
        string[] possibleNames = new[]
        {
          "ElectricalCommands.Resources.WL_Sig_Blue_Small.gif",
          "AutoCADCommands.T24Commands.Resources.WL_Sig_Blue_Small.gif",
          "Resources.WL_Sig_Blue_Small.gif"
        };

        Stream resourceStream = null;
        foreach (var name in possibleNames)
        {
          resourceStream = assembly.GetManifestResourceStream(name);
          if (resourceStream != null)
            break;
        }

        // If not found, try to find by partial name
        if (resourceStream == null)
        {
          var resourceNames = assembly.GetManifestResourceNames();
          var matchingName = resourceNames.FirstOrDefault(n => n.EndsWith(DEFAULT_SIGNATURE_FILENAME));
          if (!string.IsNullOrEmpty(matchingName))
          {
            resourceStream = assembly.GetManifestResourceStream(matchingName);
          }
        }

        if (resourceStream != null)
        {
          using (resourceStream)
          using (var fileStream = File.Create(destPath))
          {
            resourceStream.CopyTo(fileStream);
          }
          return destPath;
        }
      }
      catch
      {
        // Fall through to try file-based approach
      }

      // Try to copy from Resources folder next to DLL
      try
      {
        var assembly = Assembly.GetExecutingAssembly();
        string assemblyDir = Path.GetDirectoryName(assembly.Location);
        string resourcePath = Path.Combine(assemblyDir, "Resources", DEFAULT_SIGNATURE_FILENAME);

        if (File.Exists(resourcePath))
        {
          File.Copy(resourcePath, destPath, true);
          return destPath;
        }
      }
      catch
      {
        // Fall through
      }

      return null;
    }

    /// <summary>
    /// Validates that the selected signature file exists, falls back to default if not.
    /// </summary>
    public static SignatureEntry ValidateAndGetSignature(T24SignatureSettings settings)
    {
      var selected = GetSelectedSignature(settings);

      if (selected != null && File.Exists(selected.FilePath))
        return selected;

      // Fall back to default
      var defaultSig = settings.Signatures.FirstOrDefault(s => s.IsDefault);
      if (defaultSig != null)
      {
        settings.SelectedSignatureId = defaultSig.Id;

        // Ensure default file exists
        if (!File.Exists(defaultSig.FilePath))
        {
          string newPath = ExtractDefaultSignature();
          if (!string.IsNullOrEmpty(newPath))
          {
            defaultSig.FilePath = newPath;
          }
        }

        SaveSettings(settings);
        return defaultSig;
      }

      return null;
    }
  }
}
