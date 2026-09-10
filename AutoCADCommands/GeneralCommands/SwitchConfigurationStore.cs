using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Newtonsoft.Json;

namespace ElectricalCommands
{
  public static class SwitchConfigurationStore
  {
    private const string SettingsDirectoryName = ".acies";
    private const string SettingsFileName = "switch_config.json";
    private const string DictionaryKey = "ACIES_SWITCH_CONFIG";
    private const string XrecordKey = "SWITCH_SETTINGS_DATA";
    private const string GlobalAppName = "ProjectManagementApp";
    private const string GlobalDefaultFileName = "switch_config_default.json";

    private static readonly string[] ProjectTopLevelFolders =
    {
      "Arch",
      "Archive",
      "BGD",
      "CAD Release",
      "Documents",
      "Electrical",
      "Mechanical",
      "PDF",
      "Plumbing",
      "RFI",
      "Submittals",
      "Survey",
      "Xrefs"
    };

    private static readonly string[] DistinctiveProjectTopLevelFolders =
    {
      "Electrical",
      "Xrefs",
      "BGD",
      "CAD Release"
    };

    // In-memory cache per session
    private static ProjectSwitchSettings _sessionCache;

    public static ProjectSwitchSettings LoadSettings(Database db, Document doc = null)
    {
      if (TryResolveSettingsPath(doc, db, out string settingsPath, out _))
      {
        if (File.Exists(settingsPath))
        {
          try
          {
            string json = File.ReadAllText(settingsPath, Encoding.UTF8);
            var loaded = JsonConvert.DeserializeObject<ProjectSwitchSettings>(json);
            if (loaded != null)
            {
              loaded.EnsureDefaults();
              _sessionCache = loaded;
              return loaded;
            }
          }
          catch
          {
            // Fall back to drawing dictionary or defaults
          }
        }
      }

      // Try DWG Named Objects Dictionary
      if (db != null && TryReadFromDictionary(db, out var dictSettings))
      {
        dictSettings.EnsureDefaults();
        _sessionCache = dictSettings;
        return dictSettings;
      }

      // Try Global User Defaults
      var globalDefaults = LoadGlobalDefaults();
      if (globalDefaults != null)
      {
        _sessionCache = globalDefaults;
        return globalDefaults;
      }

      if (_sessionCache != null)
      {
        return _sessionCache;
      }

      var fresh = new ProjectSwitchSettings();
      _sessionCache = fresh;
      return fresh;
    }

    public static bool SaveSettings(Database db, ProjectSwitchSettings settings, Document doc = null)
    {
      return SaveSettings(db, settings, out _, doc);
    }

    public static bool SaveSettings(Database db, ProjectSwitchSettings settings, out string errorMessage, Document doc = null)
    {
      errorMessage = string.Empty;
      if (settings == null)
      {
        errorMessage = "Settings object is null.";
        return false;
      }

      settings.EnsureDefaults();
      settings.UpdatedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);

      string currentDwg = ResolveBestDrawingPath(doc, db);
      if (!string.IsNullOrWhiteSpace(currentDwg))
      {
        settings.SourceDrawing = currentDwg;
      }

      _sessionCache = settings;

      // 1. Save to .acies/switch_config.json
      if (TryResolveSettingsPath(doc, db, out string settingsPath, out _))
      {
        try
        {
          string directory = Path.GetDirectoryName(settingsPath);
          if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
          {
            Directory.CreateDirectory(directory);
          }

          string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
          string tempPath = settingsPath + "." + Guid.NewGuid().ToString("N") + ".tmp";

          File.WriteAllText(tempPath, json, Encoding.UTF8);
          if (File.Exists(settingsPath))
          {
            try
            {
              File.Replace(tempPath, settingsPath, null);
            }
            catch
            {
              File.Copy(tempPath, settingsPath, true);
              File.Delete(tempPath);
            }
          }
          else
          {
            File.Move(tempPath, settingsPath);
          }
        }
        catch (System.Exception ex)
        {
          errorMessage = $"Failed to save project settings file: {ex.Message}";
        }
      }

      // 2. Save to drawing dictionary
      if (db != null)
      {
        try
        {
          WriteToDictionary(db, settings);
        }
        catch (System.Exception ex)
        {
          if (string.IsNullOrEmpty(errorMessage))
          {
            errorMessage = $"Saved file, but could not update DWG dictionary: {ex.Message}";
          }
        }
      }

      return string.IsNullOrEmpty(errorMessage);
    }

    public static string GetGlobalDefaultsPath()
    {
      string userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
      if (string.IsNullOrWhiteSpace(userProfile))
      {
        userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
      }
      return Path.Combine(userProfile, "Documents", GlobalAppName, GlobalDefaultFileName);
    }

    public static ProjectSwitchSettings LoadGlobalDefaults()
    {
      try
      {
        string path = GetGlobalDefaultsPath();
        if (File.Exists(path))
        {
          string json = File.ReadAllText(path, Encoding.UTF8);
          var settings = JsonConvert.DeserializeObject<ProjectSwitchSettings>(json);
          settings?.EnsureDefaults();
          return settings;
        }
      }
      catch
      {
      }
      return null;
    }

    public static bool SaveGlobalDefaults(ProjectSwitchSettings settings, out string error)
    {
      error = string.Empty;
      try
      {
        string path = GetGlobalDefaultsPath();
        string dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
        {
          Directory.CreateDirectory(dir);
        }
        string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
        File.WriteAllText(path, json, Encoding.UTF8);
        return true;
      }
      catch (System.Exception ex)
      {
        error = ex.Message;
        return false;
      }
    }

    public static bool TryResolveSettingsPath(
      Document doc,
      Database db,
      out string settingsPath,
      out string projectRoot)
    {
      settingsPath = string.Empty;
      projectRoot = string.Empty;

      string dwgPath = ResolveBestDrawingPath(doc, db);
      if (string.IsNullOrWhiteSpace(dwgPath))
      {
        return false;
      }

      string folder = Path.GetDirectoryName(dwgPath);
      if (string.IsNullOrWhiteSpace(folder))
      {
        return false;
      }

      projectRoot = ResolveProjectRoot(folder);
      settingsPath = Path.Combine(projectRoot, SettingsDirectoryName, SettingsFileName);
      return true;
    }

    public static string ResolveBestDrawingPath(Document doc, Database db)
    {
      var candidates = new List<string>();
      if (doc != null && !string.IsNullOrWhiteSpace(doc.Name))
      {
        candidates.Add(doc.Name);
      }
      if (db != null && !string.IsNullOrWhiteSpace(db.Filename))
      {
        candidates.Add(db.Filename);
      }

      foreach (string candidate in candidates)
      {
        try
        {
          string normalized = Path.GetFullPath(candidate.Trim().Trim('"'));
          if (!Path.IsPathRooted(normalized)) continue;
          if (!string.Equals(Path.GetExtension(normalized), ".dwg", StringComparison.OrdinalIgnoreCase)) continue;
          if (File.Exists(normalized)) return normalized;
        }
        catch
        {
        }
      }
      return string.Empty;
    }

    public static string ResolveProjectRoot(string drawingFolder)
    {
      DirectoryInfo current;
      try
      {
        current = new DirectoryInfo(drawingFolder);
      }
      catch
      {
        return drawingFolder;
      }

      DirectoryInfo fallback = current;
      for (DirectoryInfo cursor = current; cursor != null; cursor = cursor.Parent)
      {
        if (DistinctiveProjectTopLevelFolders.Any(folder =>
          string.Equals(folder, cursor.Name, StringComparison.OrdinalIgnoreCase)))
        {
          return cursor.Parent?.FullName ?? cursor.FullName;
        }

        int markerCount = 0;
        foreach (string folder in ProjectTopLevelFolders)
        {
          try
          {
            if (Directory.Exists(Path.Combine(cursor.FullName, folder)))
            {
              markerCount++;
            }
          }
          catch
          {
          }
        }

        if (markerCount >= 2)
        {
          return cursor.FullName;
        }
      }

      return fallback.FullName;
    }

    private static bool TryReadFromDictionary(Database db, out ProjectSwitchSettings settings)
    {
      settings = null;
      try
      {
        using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
        {
          var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
          if (!nod.Contains(DictionaryKey)) return false;

          var settingsDict = (DBDictionary)tr.GetObject(nod.GetAt(DictionaryKey), OpenMode.ForRead);
          if (!settingsDict.Contains(XrecordKey)) return false;

          var xrec = (Xrecord)tr.GetObject(settingsDict.GetAt(XrecordKey), OpenMode.ForRead);
          var data = xrec?.Data?.AsArray();
          if (data == null || data.Length == 0) return false;

          var sb = new StringBuilder();
          foreach (var tv in data)
          {
            if (tv.TypeCode == (short)DxfCode.Text)
            {
              sb.Append(Convert.ToString(tv.Value));
            }
          }

          string json = sb.ToString();
          if (string.IsNullOrWhiteSpace(json)) return false;

          settings = JsonConvert.DeserializeObject<ProjectSwitchSettings>(json);
          return settings != null;
        }
      }
      catch
      {
        return false;
      }
    }

    private static void WriteToDictionary(Database db, ProjectSwitchSettings settings)
    {
      string json = JsonConvert.SerializeObject(settings);
      // Split into chunks of 255 chars for DXF text limit safety
      var chunks = new List<TypedValue>();
      chunks.Add(new TypedValue((int)DxfCode.Int32, settings.Version));
      const int chunkSize = 250;
      for (int i = 0; i < json.Length; i += chunkSize)
      {
        int length = Math.Min(chunkSize, json.Length - i);
        chunks.Add(new TypedValue((int)DxfCode.Text, json.Substring(i, length)));
      }

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);
        DBDictionary settingsDict;
        if (nod.Contains(DictionaryKey))
        {
          settingsDict = (DBDictionary)tr.GetObject(nod.GetAt(DictionaryKey), OpenMode.ForWrite);
        }
        else
        {
          settingsDict = new DBDictionary();
          nod.SetAt(DictionaryKey, settingsDict);
          tr.AddNewlyCreatedDBObject(settingsDict, true);
        }

        Xrecord xrec;
        if (settingsDict.Contains(XrecordKey))
        {
          xrec = (Xrecord)tr.GetObject(settingsDict.GetAt(XrecordKey), OpenMode.ForWrite);
        }
        else
        {
          xrec = new Xrecord();
          settingsDict.SetAt(XrecordKey, xrec);
          tr.AddNewlyCreatedDBObject(xrec, true);
        }

        xrec.Data = new ResultBuffer(chunks.ToArray());
        tr.Commit();
      }
    }

    /// <summary>
    /// Automatically derives North, East, South, and West orientations from a single base orientation sample.
    /// In AutoCAD plan view:
    /// North = 0 deg delta (reference)
    /// East = -90 deg (-Pi/2) rotation
    /// South = 180 deg (Pi) rotation
    /// West = +90 deg (Pi/2) rotation
    /// Text offsets are rotated around (0,0) by the orientation delta, and text rotation is kept right-reading.
    /// </summary>
    public static void AutoDeriveOrientations(
      SwitchOrientationConfig source,
      SwitchOrientation sourceOrientation,
      out SwitchOrientationConfig north,
      out SwitchOrientationConfig east,
      out SwitchOrientationConfig south,
      out SwitchOrientationConfig west)
    {
      double sourceBaseAngle = GetOrientationAngle(sourceOrientation);

      north = DeriveOrientation(source, sourceBaseAngle, GetOrientationAngle(SwitchOrientation.North));
      east = DeriveOrientation(source, sourceBaseAngle, GetOrientationAngle(SwitchOrientation.East));
      south = DeriveOrientation(source, sourceBaseAngle, GetOrientationAngle(SwitchOrientation.South));
      west = DeriveOrientation(source, sourceBaseAngle, GetOrientationAngle(SwitchOrientation.West));
    }

    public static double GetOrientationAngle(SwitchOrientation orientation)
    {
      switch (orientation)
      {
        case SwitchOrientation.North: return 0.0;
        case SwitchOrientation.East: return -Math.PI / 2.0; // 90 deg clockwise (270 deg)
        case SwitchOrientation.South: return Math.PI;       // 180 deg
        case SwitchOrientation.West: return Math.PI / 2.0;  // 90 deg counterclockwise
        default: return 0.0;
      }
    }

    private static SwitchOrientationConfig DeriveOrientation(
      SwitchOrientationConfig source,
      double fromAngle,
      double toAngle)
    {
      double deltaAngle = NormalizeAngle(toAngle - fromAngle);
      var config = source.Clone();

      if (config.Block != null)
      {
        config.Block.Rotation = NormalizeAngle(config.Block.Rotation + deltaAngle);
      }

      if (config.TextObjects != null)
      {
        foreach (var txt in config.TextObjects)
        {
          // Rotate relative offset vector by deltaAngle around (0,0)
          double cos = Math.Cos(deltaAngle);
          double sin = Math.Sin(deltaAngle);
          double x = txt.RelativeOffset.X;
          double y = txt.RelativeOffset.Y;
          txt.RelativeOffset.X = x * cos - y * sin;
          txt.RelativeOffset.Y = x * sin + y * cos;

          // Keep text right-reading: standard CAD drafting dictates text should read from bottom or right
          // Normalize text rotation so it stays between -45 deg and +135 deg (i.e. readable from bottom/right)
          double desiredRotation = NormalizeAngle(txt.Rotation + deltaAngle);
          // If upside down (between 90 deg and 270 deg), rotate 180 deg
          if (desiredRotation > Math.PI * 0.51 && desiredRotation < Math.PI * 1.51)
          {
            desiredRotation = NormalizeAngle(desiredRotation + Math.PI);
          }
          txt.Rotation = desiredRotation;
        }
      }

      return config;
    }

    public static double NormalizeAngle(double angle)
    {
      const double twoPi = 2.0 * Math.PI;
      angle = angle % twoPi;
      if (angle < 0.0) angle += twoPi;
      return angle;
    }

    /// <summary>
    /// Attempts to import a missing BlockTableRecord from a source drawing into the target database.
    /// </summary>
    public static bool TryImportBlockDefinition(
      Database targetDb,
      string sourceDwgPath,
      string blockName,
      out string errorMessage)
    {
      errorMessage = string.Empty;
      if (string.IsNullOrWhiteSpace(sourceDwgPath) || !File.Exists(sourceDwgPath))
      {
        errorMessage = $"Source drawing not found: '{sourceDwgPath}'";
        return false;
      }

      try
      {
        using (var sourceDb = new Database(false, true))
        {
          sourceDb.ReadDwgFile(sourceDwgPath, FileOpenMode.OpenForReadAndAllShare, true, "");
          ObjectId sourceBlockId = ObjectId.Null;

          using (Transaction tr = sourceDb.TransactionManager.StartTransaction())
          {
            var bt = (BlockTable)tr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
            if (bt.Has(blockName))
            {
              sourceBlockId = bt[blockName];
            }
            tr.Commit();
          }

          if (sourceBlockId.IsNull)
          {
            errorMessage = $"Block '{blockName}' does not exist in source drawing '{sourceDwgPath}'.";
            return false;
          }

          var ids = new ObjectIdCollection { sourceBlockId };
          var mapping = new IdMapping();
          sourceDb.WblockCloneObjects(ids, targetDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
          return true;
        }
      }
      catch (System.Exception ex)
      {
        errorMessage = $"Failed to clone block definition from source drawing: {ex.Message}";
        return false;
      }
    }
  }
}
