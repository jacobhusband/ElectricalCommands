using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  internal static class ProjectChecklistStore
  {
    private const string AppDataFolderName = "ProjectManagementApp";
    private const string DatabaseFileName = "project_checklists.db";
    private const string ChecklistsFileName = "checklists.json";
    private const string UpdatedByAutoCAD = "autocad";
    private static readonly Regex ProjectSegmentRegex =
      new Regex(@"^(\d{5,})\s*(?:[-_]\s*)?(.+)$", RegexOptions.Compiled);

    internal static string ResolveAppDataFolder()
    {
      string appFolder = Path.Combine(ResolveUserProfileDocumentsFolder(), AppDataFolderName);
      Directory.CreateDirectory(appFolder);
      MigrateLegacyProjectChecklistDatabase(appFolder);
      return appFolder;
    }

    internal static string ResolveDatabasePath()
    {
      return Path.Combine(ResolveAppDataFolder(), DatabaseFileName);
    }

    internal static string ResolveChecklistDefinitionsPath()
    {
      return Path.Combine(ResolveAppDataFolder(), ChecklistsFileName);
    }

    internal static List<ProjectChecklistDefinition> LoadChecklistDefinitions()
    {
      return LoadChecklistDefinitionsResult().Definitions;
    }

    internal static ProjectChecklistDefinitionsLoadResult LoadChecklistDefinitionsResult()
    {
      string path = ResolveChecklistDefinitionsPath();
      if (!File.Exists(path))
      {
        return ProjectChecklistDefinitionsLoadResult.Empty(
          path,
          $"Checklist definitions file was not found at {path}."
        );
      }

      ProjectChecklistDefinitionsFile payload;
      try
      {
        string json = File.ReadAllText(path, Encoding.UTF8);
        payload = JsonConvert.DeserializeObject<ProjectChecklistDefinitionsFile>(json)
          ?? new ProjectChecklistDefinitionsFile();
      }
      catch (System.Exception ex)
      {
        return ProjectChecklistDefinitionsLoadResult.Empty(
          path,
          $"Checklist definitions file could not be read at {path}: {ex.Message}"
        );
      }

      var definitions = new List<ProjectChecklistDefinition>();
      var seenChecklistIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      foreach (ProjectChecklistDefinition rawDefinition in payload.Checklists ?? new List<ProjectChecklistDefinition>())
      {
        ProjectChecklistDefinition definition = NormalizeDefinition(rawDefinition);
        if (string.IsNullOrWhiteSpace(definition.Id) || seenChecklistIds.Contains(definition.Id))
        {
          continue;
        }

        seenChecklistIds.Add(definition.Id);
        definitions.Add(definition);
      }

      if ((payload.Checklists ?? new List<ProjectChecklistDefinition>()).Count == 0)
      {
        return ProjectChecklistDefinitionsLoadResult.Empty(
          path,
          $"Checklist definitions file exists at {path}, but it contains zero checklists."
        );
      }

      if (definitions.Count == 0)
      {
        return ProjectChecklistDefinitionsLoadResult.Empty(
          path,
          $"Checklist definitions file exists at {path}, but no valid checklist IDs were found."
        );
      }

      return new ProjectChecklistDefinitionsLoadResult
      {
        Definitions = definitions,
        DefinitionsPath = path,
        StatusMessage = $"Loaded {definitions.Count} checklist definition(s) from {path}.",
      };
    }

    internal static ProjectChecklistStoreRecord GetRecord(string projectId)
    {
      string resolvedProjectId = NormalizePlainText(projectId);
      if (string.IsNullOrWhiteSpace(resolvedProjectId))
      {
        return null;
      }

      using (ProjectChecklistSqliteConnection db = OpenStore())
      using (ProjectChecklistSqliteStatement stmt = db.Prepare(
        @"SELECT project_id, state_json, version, updated_at_utc, updated_by
          FROM project_checklist_records
          WHERE project_id = ?1"
      ))
      {
        stmt.BindText(1, resolvedProjectId);
        return stmt.StepRow() ? ReadRecord(stmt) : null;
      }
    }

    internal static ProjectChecklistStoreRecord EnsureRecord(string projectId, string dwgPath)
    {
      ProjectChecklistStoreRecord record = GetRecord(projectId);
      if (record != null)
      {
        if (!string.IsNullOrWhiteSpace(dwgPath))
        {
          SaveLink(projectId, dwgPath);
        }
        return record;
      }

      return SaveRecord(
        projectId,
        new ProjectChecklistState(),
        dwgPath,
        UpdatedByAutoCAD,
        null
      );
    }

    internal static ProjectChecklistStoreRecord SaveRecord(
      string projectId,
      ProjectChecklistState state,
      string dwgPath,
      string updatedBy,
      long? expectedVersion
    )
    {
      string resolvedProjectId = NormalizePlainText(projectId);
      if (string.IsNullOrWhiteSpace(resolvedProjectId))
      {
        throw new InvalidOperationException("Project ID is required.");
      }

      ProjectChecklistState normalizedState = NormalizeState(state);
      string normalizedDwgPath = NormalizePath(dwgPath);
      string updatedByValue = NormalizePlainText(updatedBy);
      if (string.IsNullOrWhiteSpace(updatedByValue))
      {
        updatedByValue = UpdatedByAutoCAD;
      }

      string nowIso = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
      string stateJson = JsonConvert.SerializeObject(normalizedState);
      long previousVersion = 0;
      long nextVersion = 1;

      using (ProjectChecklistSqliteConnection db = OpenStore())
      {
        db.Exec("BEGIN IMMEDIATE;");
        try
        {
          using (ProjectChecklistSqliteStatement versionStmt = db.Prepare(
            @"SELECT version
              FROM project_checklist_records
              WHERE project_id = ?1"
          ))
          {
            versionStmt.BindText(1, resolvedProjectId);
            previousVersion = versionStmt.StepRow() ? versionStmt.GetInt64(0) : 0;
            nextVersion = previousVersion + 1;
          }

          using (ProjectChecklistSqliteStatement stmt = db.Prepare(
            @"INSERT INTO project_checklist_records (
                project_id, state_json, version, updated_at_utc, updated_by
              )
              VALUES (?1, ?2, ?3, ?4, ?5)
              ON CONFLICT(project_id) DO UPDATE SET
                state_json = excluded.state_json,
                version = excluded.version,
                updated_at_utc = excluded.updated_at_utc,
                updated_by = excluded.updated_by"
          ))
          {
            stmt.BindText(1, resolvedProjectId);
            stmt.BindText(2, stateJson);
            stmt.BindInt64(3, nextVersion);
            stmt.BindText(4, nowIso);
            stmt.BindText(5, updatedByValue);
            stmt.StepDone();
          }

          if (!string.IsNullOrWhiteSpace(normalizedDwgPath))
          {
            UpsertLink(db, resolvedProjectId, normalizedDwgPath);
          }

          db.Exec("COMMIT;");
        }
        catch
        {
          try
          {
            db.Exec("ROLLBACK;");
          }
          catch
          {
          }
          throw;
        }
      }

      ProjectChecklistStoreRecord savedRecord = GetRecord(resolvedProjectId);
      if (savedRecord != null && expectedVersion.HasValue)
      {
        savedRecord.HadConflict = expectedVersion.Value != previousVersion;
        savedRecord.PreviousVersion = previousVersion;
      }
      return savedRecord;
    }

    internal static ProjectChecklistStoreLink GetLinkByDwgPath(string dwgPath)
    {
      string normalizedDwgPath = NormalizePath(dwgPath);
      if (string.IsNullOrWhiteSpace(normalizedDwgPath))
      {
        return null;
      }

      using (ProjectChecklistSqliteConnection db = OpenStore())
      using (ProjectChecklistSqliteStatement stmt = db.Prepare(
        @"SELECT dwg_path, project_id, last_seen_at_utc
          FROM project_checklist_links
          WHERE dwg_path = ?1"
      ))
      {
        stmt.BindText(1, normalizedDwgPath);
        return stmt.StepRow() ? ReadLink(stmt) : null;
      }
    }

    internal static ProjectChecklistStoreLink SaveLink(string projectId, string dwgPath)
    {
      string resolvedProjectId = NormalizePlainText(projectId);
      string normalizedDwgPath = NormalizePath(dwgPath);
      if (string.IsNullOrWhiteSpace(resolvedProjectId) || string.IsNullOrWhiteSpace(normalizedDwgPath))
      {
        return null;
      }

      if (GetRecord(resolvedProjectId) == null)
      {
        SaveRecord(
          resolvedProjectId,
          new ProjectChecklistState(),
          string.Empty,
          "link",
          null
        );
      }

      using (ProjectChecklistSqliteConnection db = OpenStore())
      {
        db.Exec("BEGIN IMMEDIATE;");
        try
        {
          UpsertLink(db, resolvedProjectId, normalizedDwgPath);
          db.Exec("COMMIT;");
        }
        catch
        {
          try
          {
            db.Exec("ROLLBACK;");
          }
          catch
          {
          }
          throw;
        }
      }

      return GetLinkByDwgPath(normalizedDwgPath);
    }

    internal static string RequireSavedDwgPath(Document doc, Database db)
    {
      string resolvedPath = ResolveBestDwgPath(doc, db);
      if (string.IsNullOrWhiteSpace(resolvedPath))
      {
        throw new InvalidOperationException(
          "Active drawing must be saved before running this command."
        );
      }
      return resolvedPath;
    }

    internal static string ResolveBestDwgPath(Document doc, Database db)
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

      return candidates
        .Select(NormalizePath)
        .Where(path => !string.IsNullOrWhiteSpace(path))
        .Where(path => string.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase))
        .FirstOrDefault(path => !IsTemporaryPath(path))
        ?? candidates
          .Select(NormalizePath)
          .Where(path => !string.IsNullOrWhiteSpace(path))
          .FirstOrDefault(path => string.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase))
        ?? string.Empty;
    }

    internal static string ResolveProjectIdForDrawing(string dwgPath)
    {
      ProjectChecklistStoreLink link = GetLinkByDwgPath(dwgPath);
      if (!string.IsNullOrWhiteSpace(link?.ProjectId))
      {
        return link.ProjectId;
      }

      return ExtractProjectIdFromPath(dwgPath);
    }

    internal static string ExtractProjectIdFromPath(string rawPath)
    {
      if (string.IsNullOrWhiteSpace(rawPath))
      {
        return string.Empty;
      }

      string[] parts = rawPath
        .Replace("/", "\\")
        .Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
      for (int index = parts.Length - 1; index >= 0; index--)
      {
        Match match = ProjectSegmentRegex.Match(parts[index].Trim());
        if (match.Success)
        {
          return match.Groups[1].Value;
        }
      }

      return string.Empty;
    }

    internal static ProjectChecklistState NormalizeState(ProjectChecklistState state)
    {
      var normalized = new ProjectChecklistState();
      var seenChecklistIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      foreach (ProjectChecklistStateEntry entry in state?.Checklists ?? new List<ProjectChecklistStateEntry>())
      {
        string checklistId = NormalizePlainText(entry?.ChecklistId);
        if (string.IsNullOrWhiteSpace(checklistId) || seenChecklistIds.Contains(checklistId))
        {
          continue;
        }

        seenChecklistIds.Add(checklistId);
        normalized.Checklists.Add(NormalizeStateEntry(entry, checklistId));
      }

      return normalized;
    }

    internal static ProjectChecklistStateEntry NormalizeStateEntry(
      ProjectChecklistStateEntry entry,
      string checklistId
    )
    {
      var normalized = new ProjectChecklistStateEntry
      {
        ChecklistId = NormalizePlainText(checklistId),
      };

      var seenItemIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      foreach (string rawItemId in entry?.CompletedItems ?? new List<string>())
      {
        string itemId = NormalizePlainText(rawItemId);
        if (string.IsNullOrWhiteSpace(itemId) || seenItemIds.Contains(itemId))
        {
          continue;
        }

        seenItemIds.Add(itemId);
        normalized.CompletedItems.Add(itemId);
      }

      foreach (KeyValuePair<string, string> note in entry?.ItemNotes ?? new Dictionary<string, string>())
      {
        string itemId = NormalizePlainText(note.Key);
        if (string.IsNullOrWhiteSpace(itemId))
        {
          continue;
        }

        normalized.ItemNotes[itemId] = note.Value ?? string.Empty;
      }

      return normalized;
    }

    internal static string NormalizePlainText(string value)
    {
      return (value ?? string.Empty).Trim();
    }

    private static ProjectChecklistDefinition NormalizeDefinition(ProjectChecklistDefinition rawDefinition)
    {
      var definition = new ProjectChecklistDefinition
      {
        Id = NormalizePlainText(rawDefinition?.Id),
        Name = NormalizePlainText(rawDefinition?.Name),
        TemplateKey = NormalizePlainText(rawDefinition?.TemplateKey),
      };
      if (string.IsNullOrWhiteSpace(definition.Name))
      {
        definition.Name = string.IsNullOrWhiteSpace(definition.Id)
          ? "Untitled checklist"
          : definition.Id;
      }

      var seenItemIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      int fallbackIndex = 0;
      foreach (ProjectChecklistDefinitionItem rawItem in rawDefinition?.Items ?? new List<ProjectChecklistDefinitionItem>())
      {
        fallbackIndex++;
        string itemId = NormalizePlainText(rawItem?.Id);
        if (string.IsNullOrWhiteSpace(itemId))
        {
          itemId = $"item_{fallbackIndex}";
        }
        if (seenItemIds.Contains(itemId))
        {
          continue;
        }

        seenItemIds.Add(itemId);
        definition.Items.Add(new ProjectChecklistDefinitionItem
        {
          Id = itemId,
          Text = rawItem?.Text ?? string.Empty,
          Type = NormalizeChecklistRowType(rawItem?.Type),
          Order = rawItem?.Order ?? fallbackIndex - 1,
        });
      }

      definition.Items = definition.Items
        .OrderBy(item => item.Order)
        .ThenBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
        .ToList();
      return definition;
    }

    private static string NormalizeChecklistRowType(string type)
    {
      return string.Equals(type, "subheader", StringComparison.OrdinalIgnoreCase)
        ? "subheader"
        : "item";
    }

    private static ProjectChecklistStoreRecord ReadRecord(ProjectChecklistSqliteStatement stmt)
    {
      ProjectChecklistState state;
      try
      {
        state = JsonConvert.DeserializeObject<ProjectChecklistState>(
          stmt.GetText(1) ?? string.Empty
        ) ?? new ProjectChecklistState();
      }
      catch
      {
        state = new ProjectChecklistState();
      }

      return new ProjectChecklistStoreRecord
      {
        ProjectId = stmt.GetText(0),
        State = NormalizeState(state),
        Version = stmt.GetInt64(2),
        UpdatedAtUtc = stmt.GetText(3),
        UpdatedBy = stmt.GetText(4),
      };
    }

    private static ProjectChecklistStoreLink ReadLink(ProjectChecklistSqliteStatement stmt)
    {
      return new ProjectChecklistStoreLink
      {
        DwgPath = stmt.GetText(0),
        ProjectId = stmt.GetText(1),
        LastSeenAtUtc = stmt.GetText(2),
      };
    }

    private static void UpsertLink(
      ProjectChecklistSqliteConnection db,
      string projectId,
      string dwgPath
    )
    {
      using (ProjectChecklistSqliteStatement stmt = db.Prepare(
        @"INSERT INTO project_checklist_links (
            dwg_path, project_id, last_seen_at_utc
          )
          VALUES (?1, ?2, ?3)
          ON CONFLICT(dwg_path) DO UPDATE SET
            project_id = excluded.project_id,
            last_seen_at_utc = excluded.last_seen_at_utc"
      ))
      {
        stmt.BindText(1, dwgPath);
        stmt.BindText(2, projectId);
        stmt.BindText(3, DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
        stmt.StepDone();
      }
    }

    private static ProjectChecklistSqliteConnection OpenStore()
    {
      string dbPath = ResolveDatabasePath();
      ProjectChecklistSqliteConnection db = ProjectChecklistSqliteConnection.Open(dbPath);
      db.Exec("PRAGMA journal_mode = WAL;");
      db.Exec("PRAGMA foreign_keys = ON;");
      db.Exec(
        @"
CREATE TABLE IF NOT EXISTS project_checklist_records (
  project_id TEXT PRIMARY KEY,
  state_json TEXT NOT NULL,
  version INTEGER NOT NULL DEFAULT 1,
  updated_at_utc TEXT NOT NULL,
  updated_by TEXT NOT NULL DEFAULT 'unknown'
);

CREATE TABLE IF NOT EXISTS project_checklist_links (
  dwg_path TEXT PRIMARY KEY,
  project_id TEXT NOT NULL,
  last_seen_at_utc TEXT NOT NULL,
  FOREIGN KEY(project_id) REFERENCES project_checklist_records(project_id) ON DELETE CASCADE
);
"
      );
      return db;
    }

    private static string ResolveUserProfileDocumentsFolder()
    {
      string userProfile = NormalizePlainText(Environment.GetEnvironmentVariable("USERPROFILE"));
      if (string.IsNullOrWhiteSpace(userProfile))
      {
        userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
      }
      if (string.IsNullOrWhiteSpace(userProfile))
      {
        userProfile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
      }

      return Path.Combine(userProfile, "Documents");
    }

    private static string ResolveLegacyAutoCadAppDataFolder()
    {
      string legacyDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
      if (string.IsNullOrWhiteSpace(legacyDocumentsPath))
      {
        return string.Empty;
      }

      return Path.Combine(legacyDocumentsPath, AppDataFolderName);
    }

    private static void MigrateLegacyProjectChecklistDatabase(string canonicalAppFolder)
    {
      string canonicalDbPath = Path.Combine(canonicalAppFolder, DatabaseFileName);
      if (File.Exists(canonicalDbPath))
      {
        return;
      }

      string legacyFolder = ResolveLegacyAutoCadAppDataFolder();
      if (string.IsNullOrWhiteSpace(legacyFolder))
      {
        return;
      }

      string legacyDbPath = Path.Combine(legacyFolder, DatabaseFileName);
      if (
        string.Equals(canonicalDbPath, legacyDbPath, StringComparison.OrdinalIgnoreCase) ||
        !File.Exists(legacyDbPath)
      )
      {
        return;
      }

      try
      {
        File.Copy(legacyDbPath, canonicalDbPath, overwrite: false);
      }
      catch
      {
        // The canonical store can be recreated if migration fails.
      }
    }

    private static string NormalizePath(string path)
    {
      string trimmed = NormalizePlainText(path).Trim('"').Trim('\'');
      if (string.IsNullOrWhiteSpace(trimmed))
      {
        return string.Empty;
      }

      try
      {
        return Path.GetFullPath(trimmed);
      }
      catch
      {
        return trimmed;
      }
    }

    private static bool IsTemporaryPath(string path)
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        return true;
      }

      string tempPath = NormalizePath(Path.GetTempPath());
      string normalized = NormalizePath(path);
      return !string.IsNullOrWhiteSpace(tempPath) &&
        normalized.StartsWith(tempPath, StringComparison.OrdinalIgnoreCase);
    }
  }

  internal sealed class ProjectChecklistDefinitionsFile
  {
    [JsonProperty("checklists")]
    public List<ProjectChecklistDefinition> Checklists { get; set; } =
      new List<ProjectChecklistDefinition>();
  }

  internal sealed class ProjectChecklistDefinitionsLoadResult
  {
    public List<ProjectChecklistDefinition> Definitions { get; set; } =
      new List<ProjectChecklistDefinition>();
    public string DefinitionsPath { get; set; }
    public string StatusMessage { get; set; }

    internal static ProjectChecklistDefinitionsLoadResult Empty(
      string definitionsPath,
      string statusMessage
    )
    {
      return new ProjectChecklistDefinitionsLoadResult
      {
        DefinitionsPath = definitionsPath,
        StatusMessage = statusMessage,
      };
    }
  }

  internal sealed class ProjectChecklistDefinition
  {
    [JsonProperty("id")]
    public string Id { get; set; }

    [JsonProperty("name")]
    public string Name { get; set; }

    [JsonProperty("templateKey")]
    public string TemplateKey { get; set; }

    [JsonProperty("items")]
    public List<ProjectChecklistDefinitionItem> Items { get; set; } =
      new List<ProjectChecklistDefinitionItem>();
  }

  internal sealed class ProjectChecklistDefinitionItem
  {
    [JsonProperty("id")]
    public string Id { get; set; }

    [JsonProperty("text")]
    public string Text { get; set; }

    [JsonProperty("type")]
    public string Type { get; set; }

    [JsonProperty("order")]
    public int? Order { get; set; }

    [JsonIgnore]
    public bool IsSubheader =>
      string.Equals(Type, "subheader", StringComparison.OrdinalIgnoreCase);
  }

  internal sealed class ProjectChecklistState
  {
    [JsonProperty("checklists")]
    public List<ProjectChecklistStateEntry> Checklists { get; set; } =
      new List<ProjectChecklistStateEntry>();
  }

  internal sealed class ProjectChecklistStateEntry
  {
    [JsonProperty("checklistId")]
    public string ChecklistId { get; set; }

    [JsonProperty("completedItems")]
    public List<string> CompletedItems { get; set; } = new List<string>();

    [JsonProperty("itemNotes")]
    public Dictionary<string, string> ItemNotes { get; set; } =
      new Dictionary<string, string>();
  }

  internal sealed class ProjectChecklistStoreRecord
  {
    public string ProjectId { get; set; }
    public ProjectChecklistState State { get; set; }
    public long Version { get; set; }
    public string UpdatedAtUtc { get; set; }
    public string UpdatedBy { get; set; }
    public bool HadConflict { get; set; }
    public long PreviousVersion { get; set; }
  }

  internal sealed class ProjectChecklistStoreLink
  {
    public string DwgPath { get; set; }
    public string ProjectId { get; set; }
    public string LastSeenAtUtc { get; set; }
  }

  internal sealed class ProjectChecklistSqliteConnection : IDisposable
  {
    private IntPtr _db;

    private ProjectChecklistSqliteConnection(IntPtr db)
    {
      _db = db;
    }

    internal static ProjectChecklistSqliteConnection Open(string path)
    {
      IntPtr db;
      int rc = ProjectChecklistSqliteNative.sqlite3_open_v2(
        ProjectChecklistSqliteNative.ToUtf8(path),
        out db,
        ProjectChecklistSqliteNative.SQLITE_OPEN_READWRITE |
          ProjectChecklistSqliteNative.SQLITE_OPEN_CREATE,
        IntPtr.Zero
      );
      if (rc != ProjectChecklistSqliteNative.SQLITE_OK || db == IntPtr.Zero)
      {
        string message = db != IntPtr.Zero
          ? ProjectChecklistSqliteNative.GetErrorMessage(db)
          : $"sqlite3_open_v2 failed with code {rc}.";
        throw new InvalidOperationException(message);
      }

      return new ProjectChecklistSqliteConnection(db);
    }

    internal void Exec(string sql)
    {
      IntPtr errorPtr;
      int rc = ProjectChecklistSqliteNative.sqlite3_exec(
        _db,
        ProjectChecklistSqliteNative.ToUtf8(sql),
        IntPtr.Zero,
        IntPtr.Zero,
        out errorPtr
      );
      if (rc != ProjectChecklistSqliteNative.SQLITE_OK)
      {
        string message = ProjectChecklistSqliteNative.PtrToStringUtf8(errorPtr);
        if (errorPtr != IntPtr.Zero)
        {
          ProjectChecklistSqliteNative.sqlite3_free(errorPtr);
        }
        throw new InvalidOperationException(string.IsNullOrWhiteSpace(message)
          ? ProjectChecklistSqliteNative.GetErrorMessage(_db)
          : message);
      }
    }

    internal ProjectChecklistSqliteStatement Prepare(string sql)
    {
      IntPtr stmt;
      int rc = ProjectChecklistSqliteNative.sqlite3_prepare_v2(
        _db,
        ProjectChecklistSqliteNative.ToUtf8(sql),
        -1,
        out stmt,
        IntPtr.Zero
      );
      if (rc != ProjectChecklistSqliteNative.SQLITE_OK)
      {
        throw new InvalidOperationException(ProjectChecklistSqliteNative.GetErrorMessage(_db));
      }

      return new ProjectChecklistSqliteStatement(_db, stmt);
    }

    public void Dispose()
    {
      if (_db != IntPtr.Zero)
      {
        ProjectChecklistSqliteNative.sqlite3_close(_db);
        _db = IntPtr.Zero;
      }
    }
  }

  internal sealed class ProjectChecklistSqliteStatement : IDisposable
  {
    private readonly IntPtr _db;
    private IntPtr _stmt;

    internal ProjectChecklistSqliteStatement(IntPtr db, IntPtr stmt)
    {
      _db = db;
      _stmt = stmt;
    }

    internal void BindText(int index, string value)
    {
      int rc = ProjectChecklistSqliteNative.sqlite3_bind_text(
        _stmt,
        index,
        ProjectChecklistSqliteNative.ToUtf8(value ?? string.Empty),
        -1,
        ProjectChecklistSqliteNative.SQLITE_TRANSIENT
      );
      EnsureResult(rc);
    }

    internal void BindInt64(int index, long value)
    {
      int rc = ProjectChecklistSqliteNative.sqlite3_bind_int64(_stmt, index, value);
      EnsureResult(rc);
    }

    internal bool StepRow()
    {
      int rc = ProjectChecklistSqliteNative.sqlite3_step(_stmt);
      if (rc == ProjectChecklistSqliteNative.SQLITE_ROW)
      {
        return true;
      }
      if (rc == ProjectChecklistSqliteNative.SQLITE_DONE)
      {
        return false;
      }
      EnsureResult(rc);
      return false;
    }

    internal void StepDone()
    {
      int rc = ProjectChecklistSqliteNative.sqlite3_step(_stmt);
      if (rc != ProjectChecklistSqliteNative.SQLITE_DONE)
      {
        EnsureResult(rc);
      }
    }

    internal string GetText(int ordinal)
    {
      IntPtr ptr = ProjectChecklistSqliteNative.sqlite3_column_text(_stmt, ordinal);
      return ProjectChecklistSqliteNative.PtrToStringUtf8(
        ptr,
        ProjectChecklistSqliteNative.sqlite3_column_bytes(_stmt, ordinal)
      );
    }

    internal long GetInt64(int ordinal)
    {
      return ProjectChecklistSqliteNative.sqlite3_column_int64(_stmt, ordinal);
    }

    private void EnsureResult(int rc)
    {
      if (rc == ProjectChecklistSqliteNative.SQLITE_OK)
      {
        return;
      }

      throw new InvalidOperationException(ProjectChecklistSqliteNative.GetErrorMessage(_db));
    }

    public void Dispose()
    {
      if (_stmt != IntPtr.Zero)
      {
        ProjectChecklistSqliteNative.sqlite3_finalize(_stmt);
        _stmt = IntPtr.Zero;
      }
    }
  }

  internal static class ProjectChecklistSqliteNative
  {
    internal const int SQLITE_OK = 0;
    internal const int SQLITE_ROW = 100;
    internal const int SQLITE_DONE = 101;
    internal const int SQLITE_OPEN_READWRITE = 0x00000002;
    internal const int SQLITE_OPEN_CREATE = 0x00000004;
    internal static readonly IntPtr SQLITE_TRANSIENT = new IntPtr(-1);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_open_v2(
      byte[] filename,
      out IntPtr db,
      int flags,
      IntPtr zvfs
    );

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_close(IntPtr db);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern IntPtr sqlite3_errmsg(IntPtr db);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_exec(
      IntPtr db,
      byte[] sql,
      IntPtr callback,
      IntPtr arg,
      out IntPtr errmsg
    );

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_prepare_v2(
      IntPtr db,
      byte[] sql,
      int nByte,
      out IntPtr stmt,
      IntPtr tail
    );

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_bind_text(
      IntPtr stmt,
      int index,
      byte[] value,
      int n,
      IntPtr destructor
    );

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_bind_int64(IntPtr stmt, int index, long value);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_step(IntPtr stmt);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_finalize(IntPtr stmt);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern IntPtr sqlite3_column_text(IntPtr stmt, int index);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern int sqlite3_column_bytes(IntPtr stmt, int index);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern long sqlite3_column_int64(IntPtr stmt, int index);

    [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    internal static extern void sqlite3_free(IntPtr ptr);

    internal static byte[] ToUtf8(string value)
    {
      return Encoding.UTF8.GetBytes((value ?? string.Empty) + "\0");
    }

    internal static string GetErrorMessage(IntPtr db)
    {
      return PtrToStringUtf8(sqlite3_errmsg(db));
    }

    internal static string PtrToStringUtf8(IntPtr ptr)
    {
      if (ptr == IntPtr.Zero)
      {
        return string.Empty;
      }

      int length = 0;
      while (Marshal.ReadByte(ptr, length) != 0)
      {
        length++;
      }
      return PtrToStringUtf8(ptr, length);
    }

    internal static string PtrToStringUtf8(IntPtr ptr, int length)
    {
      if (ptr == IntPtr.Zero || length <= 0)
      {
        return string.Empty;
      }

      byte[] buffer = new byte[length];
      Marshal.Copy(ptr, buffer, 0, length);
      return Encoding.UTF8.GetString(buffer);
    }
  }
}
