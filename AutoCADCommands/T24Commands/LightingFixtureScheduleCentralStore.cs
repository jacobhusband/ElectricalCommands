using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string LightingFixtureScheduleDatabaseFileName = "lighting_schedules.db";
    private const string LightingFixtureScheduleStoreUpdatedByAutoCAD = "autocad";
    private const string LightingFixtureScheduleStoreUpdatedByDesktop = "desktop";

    internal static string ResolveLightingFixtureScheduleDatabasePath()
    {
      string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
      if (string.IsNullOrWhiteSpace(documentsPath))
      {
        documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
      }

      string appFolder = Path.Combine(documentsPath, "ProjectManagementApp");
      Directory.CreateDirectory(appFolder);
      return Path.Combine(appFolder, LightingFixtureScheduleDatabaseFileName);
    }

    internal static LightingFixtureScheduleStoreRecord GetLightingFixtureScheduleStoreRecord(
      string projectId
    )
    {
      string resolvedProjectId = NormalizeLightingFixtureScheduleStoreProjectId(projectId);
      if (string.IsNullOrWhiteSpace(resolvedProjectId))
      {
        return null;
      }

      using (var db = OpenLightingFixtureScheduleStore())
      using (var stmt = db.Prepare(
        @"SELECT project_id, schedule_json, target_dwg_path, table_handle, version, updated_at_utc, updated_by
          FROM lighting_schedule_projects
          WHERE project_id = ?1"
      ))
      {
        stmt.BindText(1, resolvedProjectId);
        return stmt.StepRow() ? ReadLightingFixtureScheduleStoreRecord(stmt) : null;
      }
    }

    internal static LightingFixtureScheduleStoreRecord GetLightingFixtureScheduleStoreRecordForDwg(
      string dwgPath
    )
    {
      string normalizedDwgPath = NormalizeLightingFixtureScheduleStorePath(dwgPath);
      if (string.IsNullOrWhiteSpace(normalizedDwgPath))
      {
        return null;
      }

      using (var db = OpenLightingFixtureScheduleStore())
      using (var stmt = db.Prepare(
        @"SELECT p.project_id, p.schedule_json, p.target_dwg_path, p.table_handle, p.version, p.updated_at_utc, p.updated_by
          FROM lighting_schedule_projects p
          INNER JOIN lighting_schedule_links l ON l.project_id = p.project_id
          WHERE l.dwg_path = ?1"
      ))
      {
        stmt.BindText(1, normalizedDwgPath);
        return stmt.StepRow() ? ReadLightingFixtureScheduleStoreRecord(stmt) : null;
      }
    }

    internal static LightingFixtureScheduleStoreLink GetLightingFixtureScheduleStoreLinkByDwgPath(
      string dwgPath
    )
    {
      string normalizedDwgPath = NormalizeLightingFixtureScheduleStorePath(dwgPath);
      if (string.IsNullOrWhiteSpace(normalizedDwgPath))
      {
        return null;
      }

      using (var db = OpenLightingFixtureScheduleStore())
      using (var stmt = db.Prepare(
        @"SELECT id, project_id, dwg_path, table_handle, last_applied_version, last_seen_at_utc
          FROM lighting_schedule_links
          WHERE dwg_path = ?1"
      ))
      {
        stmt.BindText(1, normalizedDwgPath);
        return stmt.StepRow() ? ReadLightingFixtureScheduleStoreLink(stmt) : null;
      }
    }

    internal static long GetLightingFixtureScheduleStoreVersion(string projectId)
    {
      LightingFixtureScheduleStoreRecord record = GetLightingFixtureScheduleStoreRecord(projectId);
      return record?.Version ?? 0;
    }

    internal static LightingFixtureScheduleStoreRecord SaveLightingFixtureScheduleStoreRecord(
      string projectId,
      LightingFixtureScheduleSyncSchedule schedule,
      string targetDwgPath,
      string tableHandle,
      string updatedBy,
      long? expectedVersion = null,
      long? lastAppliedVersion = null
    )
    {
      string resolvedProjectId = NormalizeLightingFixtureScheduleStoreProjectId(projectId);
      if (string.IsNullOrWhiteSpace(resolvedProjectId))
      {
        throw new InvalidOperationException("Lighting schedule project ID is required.");
      }

      LightingFixtureScheduleSyncSchedule normalizedSchedule = NormalizeSyncSchedule(schedule);
      string normalizedTargetDwgPath = NormalizeLightingFixtureScheduleStorePath(targetDwgPath);
      string normalizedTableHandle = NormalizePlainText(tableHandle);
      string updatedByValue = NormalizePlainText(updatedBy);
      if (string.IsNullOrWhiteSpace(updatedByValue))
      {
        updatedByValue = LightingFixtureScheduleStoreUpdatedByAutoCAD;
      }

      LightingFixtureScheduleStoreRecord previousRecord =
        GetLightingFixtureScheduleStoreRecord(resolvedProjectId);
      long previousVersion = previousRecord?.Version ?? 0;
      long nextVersion = previousVersion + 1;
      string nowIso = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);

      if (string.IsNullOrWhiteSpace(normalizedTargetDwgPath) && previousRecord != null)
      {
        normalizedTargetDwgPath = previousRecord.TargetDwgPath ?? string.Empty;
      }
      if (string.IsNullOrWhiteSpace(normalizedTableHandle) && previousRecord != null)
      {
        normalizedTableHandle = previousRecord.TableHandle ?? string.Empty;
      }

      string scheduleJson = JsonConvert.SerializeObject(
        new LightingFixtureScheduleStoreSchedulePayload
        {
          Rows = normalizedSchedule.Rows,
          GeneralNotes = normalizedSchedule.GeneralNotes,
          Notes = normalizedSchedule.Notes,
        }
      );

      using (var db = OpenLightingFixtureScheduleStore())
      {
        db.Exec("BEGIN IMMEDIATE;");

        using (var stmt = db.Prepare(
          @"INSERT INTO lighting_schedule_projects (
              project_id, schedule_json, target_dwg_path, table_handle, version, updated_at_utc, updated_by
            )
            VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)
            ON CONFLICT(project_id) DO UPDATE SET
              schedule_json = excluded.schedule_json,
              target_dwg_path = excluded.target_dwg_path,
              table_handle = excluded.table_handle,
              version = excluded.version,
              updated_at_utc = excluded.updated_at_utc,
              updated_by = excluded.updated_by"
        ))
        {
          stmt.BindText(1, resolvedProjectId);
          stmt.BindText(2, scheduleJson);
          stmt.BindText(3, normalizedTargetDwgPath ?? string.Empty);
          stmt.BindText(4, normalizedTableHandle ?? string.Empty);
          stmt.BindInt64(5, nextVersion);
          stmt.BindText(6, nowIso);
          stmt.BindText(7, updatedByValue);
          stmt.StepDone();
        }

        if (!string.IsNullOrWhiteSpace(normalizedTargetDwgPath))
        {
          UpsertLightingFixtureScheduleStoreLink(
            db,
            resolvedProjectId,
            normalizedTargetDwgPath,
            normalizedTableHandle,
            lastAppliedVersion ?? 0
          );
        }

        db.Exec("COMMIT;");
      }

      LightingFixtureScheduleStoreRecord savedRecord =
        GetLightingFixtureScheduleStoreRecord(resolvedProjectId);
      if (savedRecord != null && expectedVersion.HasValue)
      {
        savedRecord.HadConflict = expectedVersion.Value != previousVersion;
        savedRecord.PreviousVersion = previousVersion;
      }
      return savedRecord;
    }

    internal static LightingFixtureScheduleStoreLink SaveLightingFixtureScheduleStoreLink(
      string projectId,
      string dwgPath,
      string tableHandle,
      long lastAppliedVersion
    )
    {
      string resolvedProjectId = NormalizeLightingFixtureScheduleStoreProjectId(projectId);
      string normalizedDwgPath = NormalizeLightingFixtureScheduleStorePath(dwgPath);
      if (string.IsNullOrWhiteSpace(resolvedProjectId) || string.IsNullOrWhiteSpace(normalizedDwgPath))
      {
        return null;
      }

      using (var db = OpenLightingFixtureScheduleStore())
      {
        db.Exec("BEGIN IMMEDIATE;");
        UpsertLightingFixtureScheduleStoreLink(
          db,
          resolvedProjectId,
          normalizedDwgPath,
          NormalizePlainText(tableHandle),
          lastAppliedVersion
        );
        db.Exec("COMMIT;");
      }

      return GetLightingFixtureScheduleStoreLinkByDwgPath(normalizedDwgPath);
    }

    internal static string ResolveLightingFixtureScheduleProjectIdForDrawing(
      Document doc,
      Database db,
      string dwgPath
    )
    {
      string normalizedDwgPath = NormalizeLightingFixtureScheduleStorePath(dwgPath);
      LightingFixtureScheduleStoreLink link = GetLightingFixtureScheduleStoreLinkByDwgPath(normalizedDwgPath);
      if (!string.IsNullOrWhiteSpace(link?.ProjectId))
      {
        return link.ProjectId;
      }

      string extracted = ExtractProjectIdFromPath(normalizedDwgPath);
      if (!string.IsNullOrWhiteSpace(extracted))
      {
        return extracted;
      }

      return string.Empty;
    }

    internal static ObjectId ResolveLightingFixtureScheduleTableIdByHandle(
      Database db,
      string tableHandle
    )
    {
      if (db == null || string.IsNullOrWhiteSpace(tableHandle))
      {
        return ObjectId.Null;
      }

      long handleValue;
      if (!long.TryParse(tableHandle, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out handleValue))
      {
        return ObjectId.Null;
      }

      try
      {
        return db.GetObjectId(false, new Handle(handleValue), 0);
      }
      catch
      {
        return ObjectId.Null;
      }
    }

    private static string NormalizeLightingFixtureScheduleStoreProjectId(string projectId)
    {
      return NormalizePlainText(projectId);
    }

    private static string NormalizeLightingFixtureScheduleStorePath(string path)
    {
      string trimmed = NormalizePlainText(path);
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

    private static LightingFixtureScheduleStoreRecord ReadLightingFixtureScheduleStoreRecord(
      LightingFixtureScheduleSqliteStatement stmt
    )
    {
      LightingFixtureScheduleStoreSchedulePayload payload;
      try
      {
        payload = JsonConvert.DeserializeObject<LightingFixtureScheduleStoreSchedulePayload>(
          stmt.GetText(1) ?? string.Empty
        ) ?? new LightingFixtureScheduleStoreSchedulePayload();
      }
      catch
      {
        payload = new LightingFixtureScheduleStoreSchedulePayload();
      }

      LightingFixtureScheduleSyncSchedule schedule = NormalizeSyncSchedule(
        new LightingFixtureScheduleSyncSchedule
        {
          Rows = payload.Rows ?? new List<LightingFixtureScheduleSyncRow>(),
          GeneralNotes = payload.GeneralNotes,
          Notes = payload.Notes,
        }
      );

      return new LightingFixtureScheduleStoreRecord
      {
        ProjectId = stmt.GetText(0),
        Schedule = schedule,
        TargetDwgPath = stmt.GetText(2),
        TableHandle = stmt.GetText(3),
        Version = stmt.GetInt64(4),
        UpdatedAtUtc = stmt.GetText(5),
        UpdatedBy = stmt.GetText(6),
      };
    }

    private static LightingFixtureScheduleStoreLink ReadLightingFixtureScheduleStoreLink(
      LightingFixtureScheduleSqliteStatement stmt
    )
    {
      return new LightingFixtureScheduleStoreLink
      {
        Id = stmt.GetInt64(0),
        ProjectId = stmt.GetText(1),
        DwgPath = stmt.GetText(2),
        TableHandle = stmt.GetText(3),
        LastAppliedVersion = stmt.GetInt64(4),
        LastSeenAtUtc = stmt.GetText(5),
      };
    }

    private static void UpsertLightingFixtureScheduleStoreLink(
      LightingFixtureScheduleSqliteConnection db,
      string projectId,
      string dwgPath,
      string tableHandle,
      long lastAppliedVersion
    )
    {
      using (var stmt = db.Prepare(
        @"INSERT INTO lighting_schedule_links (
            project_id, dwg_path, table_handle, last_applied_version, last_seen_at_utc
          )
          VALUES (?1, ?2, ?3, ?4, ?5)
          ON CONFLICT(dwg_path) DO UPDATE SET
            project_id = excluded.project_id,
            table_handle = excluded.table_handle,
            last_applied_version = excluded.last_applied_version,
            last_seen_at_utc = excluded.last_seen_at_utc"
      ))
      {
        stmt.BindText(1, projectId);
        stmt.BindText(2, dwgPath);
        stmt.BindText(3, tableHandle ?? string.Empty);
        stmt.BindInt64(4, lastAppliedVersion);
        stmt.BindText(5, DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
        stmt.StepDone();
      }
    }

    private static LightingFixtureScheduleSqliteConnection OpenLightingFixtureScheduleStore()
    {
      string dbPath = ResolveLightingFixtureScheduleDatabasePath();
      var db = LightingFixtureScheduleSqliteConnection.Open(dbPath);
      db.Exec("PRAGMA journal_mode = WAL;");
      db.Exec("PRAGMA foreign_keys = ON;");
      db.Exec(
        @"
CREATE TABLE IF NOT EXISTS lighting_schedule_projects (
  project_id TEXT PRIMARY KEY,
  schedule_json TEXT NOT NULL,
  target_dwg_path TEXT NOT NULL DEFAULT '',
  table_handle TEXT NOT NULL DEFAULT '',
  version INTEGER NOT NULL DEFAULT 1,
  updated_at_utc TEXT NOT NULL,
  updated_by TEXT NOT NULL DEFAULT 'unknown'
);

CREATE TABLE IF NOT EXISTS lighting_schedule_links (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id TEXT NOT NULL,
  dwg_path TEXT NOT NULL UNIQUE,
  table_handle TEXT NOT NULL DEFAULT '',
  last_applied_version INTEGER NOT NULL DEFAULT 0,
  last_seen_at_utc TEXT NOT NULL DEFAULT '',
  FOREIGN KEY(project_id) REFERENCES lighting_schedule_projects(project_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_lighting_schedule_links_project_id
  ON lighting_schedule_links(project_id);
"
      );
      return db;
    }
  }

  internal sealed class LightingFixtureScheduleStoreRecord
  {
    public string ProjectId { get; set; }
    public LightingFixtureScheduleSyncSchedule Schedule { get; set; }
    public string TargetDwgPath { get; set; }
    public string TableHandle { get; set; }
    public long Version { get; set; }
    public string UpdatedAtUtc { get; set; }
    public string UpdatedBy { get; set; }
    public bool HadConflict { get; set; }
    public long PreviousVersion { get; set; }
  }

  internal sealed class LightingFixtureScheduleStoreLink
  {
    public long Id { get; set; }
    public string ProjectId { get; set; }
    public string DwgPath { get; set; }
    public string TableHandle { get; set; }
    public long LastAppliedVersion { get; set; }
    public string LastSeenAtUtc { get; set; }
  }

  internal sealed class LightingFixtureScheduleStoreSchedulePayload
  {
    public List<LightingFixtureScheduleSyncRow> Rows { get; set; } =
      new List<LightingFixtureScheduleSyncRow>();
    public string GeneralNotes { get; set; }
    public string Notes { get; set; }
  }

  internal sealed class LightingFixtureScheduleSqliteConnection : IDisposable
  {
    private IntPtr _db;

    private LightingFixtureScheduleSqliteConnection(IntPtr db)
    {
      _db = db;
    }

    internal static LightingFixtureScheduleSqliteConnection Open(string path)
    {
      IntPtr db;
      int rc = LightingFixtureScheduleSqliteNative.sqlite3_open_v2(
        LightingFixtureScheduleSqliteNative.ToUtf8(path),
        out db,
        LightingFixtureScheduleSqliteNative.SQLITE_OPEN_READWRITE |
          LightingFixtureScheduleSqliteNative.SQLITE_OPEN_CREATE,
        IntPtr.Zero
      );
      if (rc != LightingFixtureScheduleSqliteNative.SQLITE_OK || db == IntPtr.Zero)
      {
        string message = db != IntPtr.Zero
          ? LightingFixtureScheduleSqliteNative.GetErrorMessage(db)
          : $"sqlite3_open_v2 failed with code {rc}.";
        throw new InvalidOperationException(message);
      }

      return new LightingFixtureScheduleSqliteConnection(db);
    }

    internal void Exec(string sql)
    {
      IntPtr errorPtr;
      int rc = LightingFixtureScheduleSqliteNative.sqlite3_exec(
        _db,
        LightingFixtureScheduleSqliteNative.ToUtf8(sql),
        IntPtr.Zero,
        IntPtr.Zero,
        out errorPtr
      );
      if (rc != LightingFixtureScheduleSqliteNative.SQLITE_OK)
      {
        string message = LightingFixtureScheduleSqliteNative.PtrToStringUtf8(errorPtr);
        if (errorPtr != IntPtr.Zero)
        {
          LightingFixtureScheduleSqliteNative.sqlite3_free(errorPtr);
        }
        throw new InvalidOperationException(string.IsNullOrWhiteSpace(message)
          ? LightingFixtureScheduleSqliteNative.GetErrorMessage(_db)
          : message);
      }
    }

    internal LightingFixtureScheduleSqliteStatement Prepare(string sql)
    {
      IntPtr stmt;
      int rc = LightingFixtureScheduleSqliteNative.sqlite3_prepare_v2(
        _db,
        LightingFixtureScheduleSqliteNative.ToUtf8(sql),
        -1,
        out stmt,
        IntPtr.Zero
      );
      if (rc != LightingFixtureScheduleSqliteNative.SQLITE_OK)
      {
        throw new InvalidOperationException(LightingFixtureScheduleSqliteNative.GetErrorMessage(_db));
      }

      return new LightingFixtureScheduleSqliteStatement(_db, stmt);
    }

    public void Dispose()
    {
      if (_db != IntPtr.Zero)
      {
        LightingFixtureScheduleSqliteNative.sqlite3_close(_db);
        _db = IntPtr.Zero;
      }
    }
  }

  internal sealed class LightingFixtureScheduleSqliteStatement : IDisposable
  {
    private readonly IntPtr _db;
    private IntPtr _stmt;

    internal LightingFixtureScheduleSqliteStatement(IntPtr db, IntPtr stmt)
    {
      _db = db;
      _stmt = stmt;
    }

    internal void BindText(int index, string value)
    {
      int rc = LightingFixtureScheduleSqliteNative.sqlite3_bind_text(
        _stmt,
        index,
        LightingFixtureScheduleSqliteNative.ToUtf8(value ?? string.Empty),
        -1,
        LightingFixtureScheduleSqliteNative.SQLITE_TRANSIENT
      );
      EnsureResult(rc);
    }

    internal void BindInt64(int index, long value)
    {
      int rc = LightingFixtureScheduleSqliteNative.sqlite3_bind_int64(_stmt, index, value);
      EnsureResult(rc);
    }

    internal bool StepRow()
    {
      int rc = LightingFixtureScheduleSqliteNative.sqlite3_step(_stmt);
      if (rc == LightingFixtureScheduleSqliteNative.SQLITE_ROW)
      {
        return true;
      }
      if (rc == LightingFixtureScheduleSqliteNative.SQLITE_DONE)
      {
        return false;
      }
      EnsureResult(rc);
      return false;
    }

    internal void StepDone()
    {
      int rc = LightingFixtureScheduleSqliteNative.sqlite3_step(_stmt);
      if (rc != LightingFixtureScheduleSqliteNative.SQLITE_DONE)
      {
        EnsureResult(rc);
      }
    }

    internal string GetText(int ordinal)
    {
      IntPtr ptr = LightingFixtureScheduleSqliteNative.sqlite3_column_text(_stmt, ordinal);
      return LightingFixtureScheduleSqliteNative.PtrToStringUtf8(
        ptr,
        LightingFixtureScheduleSqliteNative.sqlite3_column_bytes(_stmt, ordinal)
      );
    }

    internal long GetInt64(int ordinal)
    {
      return LightingFixtureScheduleSqliteNative.sqlite3_column_int64(_stmt, ordinal);
    }

    private void EnsureResult(int rc)
    {
      if (rc == LightingFixtureScheduleSqliteNative.SQLITE_OK)
      {
        return;
      }

      throw new InvalidOperationException(LightingFixtureScheduleSqliteNative.GetErrorMessage(_db));
    }

    public void Dispose()
    {
      if (_stmt != IntPtr.Zero)
      {
        LightingFixtureScheduleSqliteNative.sqlite3_finalize(_stmt);
        _stmt = IntPtr.Zero;
      }
    }
  }

  internal static class LightingFixtureScheduleSqliteNative
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
