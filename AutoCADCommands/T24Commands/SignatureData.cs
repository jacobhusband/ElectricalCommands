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
  /// Represents a single template entry for T24 form data.
  /// </summary>
  public class T24TemplateEntry
  {
    /// <summary>
    /// Unique identifier for the template.
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// User-friendly template name (e.g., "Wilson").
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Responsible person name (e.g., "Wilson Lee, PE").
    /// </summary>
    public string ResponsibleName { get; set; }

    /// <summary>
    /// Company name.
    /// </summary>
    public string Company { get; set; }

    /// <summary>
    /// Address line 1.
    /// </summary>
    public string AddressLine1 { get; set; }

    /// <summary>
    /// Address line 2.
    /// </summary>
    public string AddressLine2 { get; set; }

    /// <summary>
    /// Phone number.
    /// </summary>
    public string Phone { get; set; }

    /// <summary>
    /// License number.
    /// </summary>
    public string License { get; set; }

    /// <summary>
    /// Signature ID to use with this template.
    /// </summary>
    public string SignatureId { get; set; }

    /// <summary>
    /// True if this is the bundled default template (cannot be deleted).
    /// </summary>
    public bool IsDefault { get; set; }

    /// <summary>
    /// When this template was added.
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

    /// <summary>
    /// ID of the currently selected/active template.
    /// </summary>
    public string SelectedTemplateId { get; set; }

    /// <summary>
    /// List of all available templates.
    /// </summary>
    public List<T24TemplateEntry> Templates { get; set; } = new List<T24TemplateEntry>();
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
    private const string DEFAULT_TEMPLATE_NAME = "Wilson";
    private const string DEFAULT_RESPONSIBLE_NAME = "Wilson Lee, PE";
    private const string DEFAULT_COMPANY = "ACIES Engineering";
    private const string DEFAULT_ADDRESS1 = "400 N McCarthy Blvd., Suite 250";
    private const string DEFAULT_ADDRESS2 = "Milipitas, CA 95035";
    private const string DEFAULT_PHONE = "(408) 522-5255";
    private const string DEFAULT_LICENSE = "E015418";

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
      bool settingsChanged = false;

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
      if (settings == null)
      {
        settings = InitializeDefaultSettings();
        return settings;
      }

      if (settings.Signatures == null)
      {
        settings.Signatures = new List<SignatureEntry>();
        settingsChanged = true;
      }

      if (settings.Signatures.Count == 0)
      {
        // Ensure default signature exists
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
          settingsChanged = true;
        }
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
          settingsChanged = true;
        }
        else if (settings.Signatures.Count > 0)
        {
          settings.SelectedSignatureId = settings.Signatures[0].Id;
          settingsChanged = true;
        }
      }

      settingsChanged |= EnsureTemplates(settings);

      var activeTemplate = GetSelectedTemplate(settings);
      if (activeTemplate != null &&
          !string.IsNullOrEmpty(activeTemplate.SignatureId) &&
          activeTemplate.SignatureId != settings.SelectedSignatureId)
      {
        settings.SelectedSignatureId = activeTemplate.SignatureId;
        settingsChanged = true;
      }

      if (settingsChanged)
      {
        SaveSettings(settings);
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

      // Update templates referencing this signature
      var defaultSig = settings.Signatures.FirstOrDefault(s => s.IsDefault) ?? settings.Signatures.FirstOrDefault();
      if (settings.Templates != null && settings.Templates.Count > 0)
      {
        foreach (var template in settings.Templates.Where(t => t.SignatureId == signatureId))
        {
          template.SignatureId = defaultSig?.Id;
        }
      }

      // If this was the selected signature, switch to default
      if (settings.SelectedSignatureId == signatureId)
      {
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
        Signatures = new List<SignatureEntry>(),
        Templates = new List<T24TemplateEntry>()
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

      // Create default template
      var defaultTemplate = CreateDefaultTemplate(settings);
      if (defaultTemplate != null)
      {
        settings.Templates.Add(defaultTemplate);
        settings.SelectedTemplateId = defaultTemplate.Id;
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

    /// <summary>
    /// Gets the currently selected template entry.
    /// </summary>
    public static T24TemplateEntry GetSelectedTemplate(T24SignatureSettings settings)
    {
      if (settings == null || settings.Templates == null || settings.Templates.Count == 0)
        return null;

      if (string.IsNullOrEmpty(settings.SelectedTemplateId))
        return settings.Templates.FirstOrDefault();

      return settings.Templates.FirstOrDefault(t => t.Id == settings.SelectedTemplateId) ?? settings.Templates.FirstOrDefault();
    }

    /// <summary>
    /// Gets the signature entry associated with a template.
    /// </summary>
    public static SignatureEntry GetSignatureForTemplate(T24SignatureSettings settings, T24TemplateEntry template)
    {
      if (settings == null || template == null || string.IsNullOrEmpty(template.SignatureId))
        return null;

      return settings.Signatures.FirstOrDefault(s => s.Id == template.SignatureId);
    }

    /// <summary>
    /// Validates selected template and signature; falls back to defaults if needed.
    /// </summary>
    public static T24TemplateEntry ValidateAndGetTemplate(T24SignatureSettings settings)
    {
      if (settings == null)
        return null;

      bool settingsChanged = false;

      if (settings.Templates == null || settings.Templates.Count == 0)
      {
        settingsChanged |= EnsureTemplates(settings);
      }

      var template = GetSelectedTemplate(settings);
      if (template == null)
      {
        settingsChanged |= EnsureTemplates(settings);
        template = GetSelectedTemplate(settings);
      }

      if (template != null)
      {
        var signature = GetSignatureForTemplate(settings, template);
        if (signature == null)
        {
          var defaultSig = settings.Signatures.FirstOrDefault(s => s.IsDefault) ?? settings.Signatures.FirstOrDefault();
          template.SignatureId = defaultSig?.Id;
          signature = defaultSig;
          settingsChanged = true;
        }

        if (signature != null && !File.Exists(signature.FilePath))
        {
          signature = ValidateAndGetSignature(settings);
          if (signature != null)
          {
            template.SignatureId = signature.Id;
            settingsChanged = true;
          }
        }
      }

      if (settingsChanged)
      {
        SaveSettings(settings);
      }

      return template;
    }

    /// <summary>
    /// Adds a new template, optionally cloning an existing template.
    /// </summary>
    public static T24TemplateEntry AddTemplate(T24SignatureSettings settings, T24TemplateEntry sourceTemplate, string name)
    {
      if (settings == null || string.IsNullOrWhiteSpace(name))
        return null;

      var template = new T24TemplateEntry
      {
        Id = Guid.NewGuid().ToString(),
        Name = name.Trim(),
        ResponsibleName = sourceTemplate?.ResponsibleName ?? DEFAULT_RESPONSIBLE_NAME,
        Company = sourceTemplate?.Company ?? DEFAULT_COMPANY,
        AddressLine1 = sourceTemplate?.AddressLine1 ?? DEFAULT_ADDRESS1,
        AddressLine2 = sourceTemplate?.AddressLine2 ?? DEFAULT_ADDRESS2,
        Phone = sourceTemplate?.Phone ?? DEFAULT_PHONE,
        License = sourceTemplate?.License ?? DEFAULT_LICENSE,
        SignatureId = sourceTemplate?.SignatureId ?? settings.SelectedSignatureId ?? settings.Signatures.FirstOrDefault()?.Id,
        IsDefault = false,
        DateAdded = DateTime.Now
      };

      settings.Templates.Add(template);
      return template;
    }

    /// <summary>
    /// Renames a template.
    /// </summary>
    public static bool RenameTemplate(T24SignatureSettings settings, string templateId, string newName)
    {
      var entry = settings?.Templates?.FirstOrDefault(t => t.Id == templateId);
      if (entry == null || string.IsNullOrWhiteSpace(newName))
        return false;

      entry.Name = newName.Trim();
      return true;
    }

    /// <summary>
    /// Removes a template (cannot remove default).
    /// </summary>
    public static bool RemoveTemplate(T24SignatureSettings settings, string templateId)
    {
      var entry = settings?.Templates?.FirstOrDefault(t => t.Id == templateId);
      if (entry == null || entry.IsDefault)
        return false;

      settings.Templates.Remove(entry);

      if (settings.SelectedTemplateId == templateId)
      {
        var defaultTemplate = settings.Templates.FirstOrDefault(t => t.IsDefault);
        settings.SelectedTemplateId = defaultTemplate?.Id ?? settings.Templates.FirstOrDefault()?.Id;
      }

      return true;
    }

    private static T24TemplateEntry CreateDefaultTemplate(T24SignatureSettings settings)
    {
      var selectedSig = settings.Signatures.FirstOrDefault(s => s.Id == settings.SelectedSignatureId);
      var defaultSig = selectedSig ?? settings.Signatures.FirstOrDefault(s => s.IsDefault) ?? settings.Signatures.FirstOrDefault();

      return new T24TemplateEntry
      {
        Id = Guid.NewGuid().ToString(),
        Name = DEFAULT_TEMPLATE_NAME,
        ResponsibleName = DEFAULT_RESPONSIBLE_NAME,
        Company = DEFAULT_COMPANY,
        AddressLine1 = DEFAULT_ADDRESS1,
        AddressLine2 = DEFAULT_ADDRESS2,
        Phone = DEFAULT_PHONE,
        License = DEFAULT_LICENSE,
        SignatureId = defaultSig?.Id,
        IsDefault = true,
        DateAdded = DateTime.Now
      };
    }

    private static bool EnsureTemplates(T24SignatureSettings settings)
    {
      bool settingsChanged = false;

      if (settings.Templates == null)
      {
        settings.Templates = new List<T24TemplateEntry>();
        settingsChanged = true;
      }

      if (settings.Templates.Count == 0)
      {
        var defaultTemplate = CreateDefaultTemplate(settings);
        if (defaultTemplate != null)
        {
          settings.Templates.Add(defaultTemplate);
          settings.SelectedTemplateId = defaultTemplate.Id;
          settingsChanged = true;
        }
      }

      if (string.IsNullOrEmpty(settings.SelectedTemplateId) ||
          !settings.Templates.Any(t => t.Id == settings.SelectedTemplateId))
      {
        var defaultTemplate = settings.Templates.FirstOrDefault(t => t.IsDefault);
        settings.SelectedTemplateId = defaultTemplate?.Id ?? settings.Templates.FirstOrDefault()?.Id;
        settingsChanged = true;
      }

      if (!settings.Templates.Any(t => t.IsDefault) && settings.Templates.Count > 0)
      {
        settings.Templates[0].IsDefault = true;
        settingsChanged = true;
      }

      var defaultSig = settings.Signatures.FirstOrDefault(s => s.IsDefault) ?? settings.Signatures.FirstOrDefault();
      foreach (var template in settings.Templates)
      {
        if (string.IsNullOrEmpty(template.SignatureId) ||
            !settings.Signatures.Any(s => s.Id == template.SignatureId))
        {
          template.SignatureId = defaultSig?.Id;
          settingsChanged = true;
        }
      }

      return settingsChanged;
    }
  }
}
