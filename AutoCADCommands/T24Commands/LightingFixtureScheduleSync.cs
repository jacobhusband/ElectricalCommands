using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string LightingFixtureScheduleSyncFileName = "T24LightingFixtureSchedule.sync.json";
    private static readonly string[] LightingFixtureScheduleHeaders =
    {
      "MARK",
      "DESCRIPTION",
      "MANUFACTURER",
      "MODEL NUMBER",
      "MOUNTING",
      "VOLTS",
      "WATTS",
      "NOTES"
    };
    private static readonly Regex LightingProjectSegmentRegex =
      new Regex(@"^(\d{5,})\s*(?:[-_]\s*)?(.+)$", RegexOptions.Compiled);
    private static readonly Dictionary<string, string> CaliforniaStarterModelFragments =
      new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
      {
        { "L1", "IC1JBPF 07LM 30K 90CRI 120 FRPC" },
        { "L2", "6020EN3-15" },
        { "L3", "49119EN3-962" },
        { "L4", "WF4 ADJ SWW5 90CRI MW M6" },
        { "L5", "8909PEN3-12" },
        { "L6", "WF4 REG SWW5 90CRI MW M6" },
        { "L7", "T24M-2C-TUBS-SP-30K-WH" },
        { "L8", "JSBC 6IN 30K 90CRI WH" },
        { "L9", "TR24M-2C-WH" },
        { "L10", "T24M LINEA" },
        { "L11", "HTLHD-WW-16" },
      };

    internal static System.Drawing.Image LoadLightingFixtureSymbolPreview(
      string assetPath,
      string starterFixtureKey = null
    )
    {
      try
      {
        using (Stream stream = OpenLightingFixtureSymbolStream(assetPath, starterFixtureKey))
        {
          if (stream == null)
          {
            return null;
          }
          using (System.Drawing.Image source = System.Drawing.Image.FromStream(stream))
          {
            return new System.Drawing.Bitmap(source);
          }
        }
      }
      catch
      {
        return null;
      }
    }

    private static Stream OpenLightingFixtureSymbolStream(
      string assetPath,
      string starterFixtureKey = null
    )
    {
      string normalizedAssetPath = NormalizePlainText(assetPath)
        .Replace('\\', '/');
      string fileName = Path.GetFileName(normalizedAssetPath);
      if (
        string.IsNullOrWhiteSpace(fileName) &&
        !string.IsNullOrWhiteSpace(starterFixtureKey)
      )
      {
        Match match = Regex.Match(
          starterFixtureKey,
          @"ca-2025-res-l(\d+)$",
          RegexOptions.IgnoreCase
        );
        if (match.Success)
        {
          fileName = $"ca-residential-l{match.Groups[1].Value}.png";
        }
      }

      if (
        normalizedAssetPath.StartsWith("assets/lighting/", StringComparison.OrdinalIgnoreCase) ||
        (!string.IsNullOrWhiteSpace(fileName) &&
          fileName.StartsWith("ca-residential-l", StringComparison.OrdinalIgnoreCase))
      )
      {
        Assembly assembly = typeof(GeneralCommands).Assembly;
        string resourceName = assembly.GetManifestResourceNames().FirstOrDefault(name =>
          name.EndsWith(
            $"LightingSymbols.{fileName}",
            StringComparison.OrdinalIgnoreCase
          )
        );
        return string.IsNullOrWhiteSpace(resourceName)
          ? null
          : assembly.GetManifestResourceStream(resourceName);
      }

      if (normalizedAssetPath.StartsWith("page_assets/", StringComparison.OrdinalIgnoreCase))
      {
        string appDataRoot = Path.GetDirectoryName(
          ResolveLightingFixtureScheduleDatabasePath()
        );
        string relativePath = normalizedAssetPath.Replace('/', Path.DirectorySeparatorChar);
        string resolvedPath = Path.GetFullPath(Path.Combine(appDataRoot, relativePath));
        string safeRoot = Path.GetFullPath(appDataRoot)
          .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) +
          Path.DirectorySeparatorChar;
        if (
          resolvedPath.StartsWith(safeRoot, StringComparison.OrdinalIgnoreCase) &&
          File.Exists(resolvedPath)
        )
        {
          return File.Open(resolvedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
      }

      return null;
    }

    public void LightingFixtureSchedulePull()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      try
      {
        string dwgPath = RequireSavedDwgPath(doc, db);
        string projectId = ResolveLightingFixtureScheduleProjectIdForDrawing(doc, db, dwgPath);
        if (string.IsNullOrWhiteSpace(projectId))
        {
          projectId = PromptForLightingFixtureScheduleProjectId(ed);
        }
        if (string.IsNullOrWhiteSpace(projectId))
        {
          ed.WriteMessage("\nLFSPULL cancelled. Project ID is required.");
          return;
        }

        ObjectId tableId = PromptForLightingFixtureScheduleTable(
          ed,
          "\nSelect lighting fixture schedule table to export to the central store: "
        );
        if (tableId == ObjectId.Null)
        {
          ed.WriteMessage("\nLFSPULL cancelled.");
          return;
        }

        string tableHandle = "UNKNOWN";
        int rowCount = 0;
        LightingFixtureScheduleStoreRecord record = null;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          Table table = tr.GetObject(tableId, OpenMode.ForRead, false) as Table;
          if (table == null)
          {
            ed.WriteMessage("\nSelected object is not an AutoCAD table.");
            return;
          }

          LightingFixtureScheduleSyncSchedule schedule = ExtractLightingScheduleFromTable(table);
          schedule = NormalizeSyncSchedule(schedule);
          rowCount = schedule.Rows.Count;
          tableHandle = table.Handle.ToString();
          record = SaveLightingFixtureScheduleStoreRecord(
            projectId,
            schedule,
            dwgPath,
            tableHandle,
            LightingFixtureScheduleStoreUpdatedByAutoCAD,
            null,
            0
          );

          tr.Commit();
        }

        SaveLightingFixtureScheduleStoreLink(
          projectId,
          dwgPath,
          tableHandle,
          record?.Version ?? 0
        );
        ed.WriteMessage(
          $"\nLFSPULL complete. Project: {projectId}, Handle: {tableHandle}, Rows: {rowCount}, Version: {record?.Version ?? 0}, Store: {ResolveLightingFixtureScheduleDatabasePath()}"
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nLFSPULL error: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    public void LightingFixtureSchedulePush()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      try
      {
        string dwgPath = RequireSavedDwgPath(doc, db);
        string projectId = ResolveLightingFixtureScheduleProjectIdForDrawing(doc, db, dwgPath);
        if (string.IsNullOrWhiteSpace(projectId))
        {
          LightingFixtureScheduleStoreLink link = GetLightingFixtureScheduleStoreLinkByDwgPath(dwgPath);
          projectId = NormalizePlainText(link?.ProjectId);
        }
        if (string.IsNullOrWhiteSpace(projectId))
        {
          projectId = PromptForLightingFixtureScheduleProjectId(ed);
        }
        if (string.IsNullOrWhiteSpace(projectId))
        {
          ed.WriteMessage("\nLFSPUSH cancelled. Project ID is required.");
          return;
        }

        LightingFixtureScheduleStoreRecord record = GetLightingFixtureScheduleStoreRecord(projectId);
        if (record == null)
        {
          ed.WriteMessage(
            $"\nLFSPUSH cancelled. No central store record was found for project {projectId}."
          );
          return;
        }

        ObjectId tableId = PromptForLightingFixtureScheduleTable(
          ed,
          "\nSelect lighting fixture schedule table to update from the central store: "
        );
        if (tableId == ObjectId.Null)
        {
          ed.WriteMessage("\nLFSPUSH cancelled.");
          return;
        }

        int rowCount = record.Schedule?.Rows?.Count ?? 0;
        string tableHandle = "UNKNOWN";
        LightingFixtureScheduleVisualNormalizationResult normalizationResult = null;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          Table table = tr.GetObject(tableId, OpenMode.ForWrite, false) as Table;
          if (table == null)
          {
            ed.WriteMessage("\nSelected object is not an AutoCAD table.");
            return;
          }

          normalizationResult = ApplyLightingScheduleToTable(table, db, tr, record.Schedule);
          table.GenerateLayout();
          TryRecomputeTableBlock(table);
          tableHandle = table.Handle.ToString();
          tr.Commit();
        }

        SaveLightingFixtureScheduleStoreLink(projectId, dwgPath, tableHandle, record.Version);
        SafeEditorRegen(ed);
        ed.WriteMessage(
          $"\nLFSPUSH complete. Project: {projectId}, Handle: {tableHandle}, Rows applied: {rowCount}, Version: {record.Version}. Visual normalization: {FormatNormalizationSummary(normalizationResult)}"
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nLFSPUSH error: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    internal static bool TryApplyLightingFixtureScheduleSyncToTable(
      Table table,
      Database db,
      out string message,
      out bool isError
    )
    {
      message = string.Empty;
      isError = false;
      try
      {
        if (table == null || db == null)
        {
          return false;
        }

        Document activeDoc = Application.DocumentManager?.MdiActiveDocument;
        string dwgPath = ResolveBestDwgPath(activeDoc, db);
        LightingFixtureScheduleStoreRecord record = null;
        if (!string.IsNullOrWhiteSpace(dwgPath))
        {
          LightingFixtureScheduleStoreLink link =
            GetLightingFixtureScheduleStoreLinkByDwgPath(dwgPath);
          if (!string.IsNullOrWhiteSpace(link?.ProjectId))
          {
            record = GetLightingFixtureScheduleStoreRecord(link.ProjectId);
          }
          record = record ?? GetLightingFixtureScheduleStoreRecordForDwg(dwgPath);
        }

        if (record != null)
        {
          Transaction tr = db.TransactionManager.TopTransaction;
          if (tr == null)
          {
            throw new InvalidOperationException(
              "An active transaction is required to apply lighting fixture schedule sync data."
            );
          }

          LightingFixtureScheduleVisualNormalizationResult normalizationResult =
            ApplyLightingScheduleToTable(table, db, tr, record.Schedule);
          message =
            $"Applied central-store data for project {record.ProjectId}. Visual normalization: {FormatNormalizationSummary(normalizationResult)}.";
          return true;
        }

        string syncPath = ResolveLightingFixtureScheduleSyncPath(activeDoc, db);
        if (!File.Exists(syncPath))
        {
          return false;
        }

        LightingFixtureScheduleSyncPayload payload = NormalizeSyncPayload(
          ReadLightingFixtureScheduleSyncFile(syncPath)
        );
        Transaction fallbackTr = db.TransactionManager.TopTransaction;
        if (fallbackTr == null)
        {
          throw new InvalidOperationException(
            "An active transaction is required to apply lighting fixture schedule sync data."
          );
        }

        LightingFixtureScheduleVisualNormalizationResult fallbackNormalization =
          ApplyLightingScheduleToTable(table, db, fallbackTr, payload.Schedule);
        message =
          $"Applied legacy sync-file data from {syncPath}. Visual normalization: {FormatNormalizationSummary(fallbackNormalization)}.";
        return true;
      }
      catch (System.Exception ex)
      {
        isError = true;
        message = ex.Message;
        return false;
      }
    }

    private static ObjectId PromptForLightingFixtureScheduleTable(Editor ed, string prompt)
    {
      PromptEntityOptions peo = new PromptEntityOptions(prompt);
      peo.SetRejectMessage("\nSelected object is not an AutoCAD table.");
      peo.AddAllowedClass(typeof(Table), true);
      PromptEntityResult per = ed.GetEntity(peo);
      return per.Status == PromptStatus.OK ? per.ObjectId : ObjectId.Null;
    }

    private static string RequireSavedDwgPath(Document doc, Database db)
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

    private static string ResolveLightingFixtureScheduleSyncPath(Document doc, Database db)
    {
      string dwgPath = RequireSavedDwgPath(doc, db);
      string folder = Path.GetDirectoryName(dwgPath);
      if (string.IsNullOrWhiteSpace(folder))
      {
        throw new InvalidOperationException(
          "Unable to resolve drawing folder for sync file."
        );
      }
      return Path.Combine(folder, LightingFixtureScheduleSyncFileName);
    }

    private static string ResolveBestDwgPath(Document doc, Database db)
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

      if (candidates.Count == 0)
      {
        return string.Empty;
      }

      var normalizedCandidates = candidates
        .Select(NormalizePathSafely)
        .Where(path => !string.IsNullOrWhiteSpace(path))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();
      if (normalizedCandidates.Count == 0)
      {
        return string.Empty;
      }

      string preferred = normalizedCandidates
        .Where(IsLikelySavedDwgPath)
        .Where(path => !IsTemporaryPath(path))
        .FirstOrDefault();
      if (!string.IsNullOrWhiteSpace(preferred))
      {
        return preferred;
      }

      preferred = normalizedCandidates.Where(IsLikelySavedDwgPath).FirstOrDefault();
      if (!string.IsNullOrWhiteSpace(preferred))
      {
        return preferred;
      }

      return string.Empty;
    }

    private static string NormalizePathSafely(string path)
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        return string.Empty;
      }

      try
      {
        return Path.GetFullPath(path.Trim().Trim('"'));
      }
      catch
      {
        return string.Empty;
      }
    }

    private static bool IsLikelySavedDwgPath(string path)
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        return false;
      }

      if (!Path.IsPathRooted(path))
      {
        return false;
      }

      if (!string.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase))
      {
        return false;
      }

      string directory = Path.GetDirectoryName(path);
      return !string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory);
    }

    private static bool IsTemporaryPath(string path)
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        return false;
      }

      try
      {
        string tempPath = Path.GetFullPath(Path.GetTempPath()).TrimEnd('\\');
        string normalized = Path.GetFullPath(path).TrimEnd('\\');
        return normalized.StartsWith(tempPath, StringComparison.OrdinalIgnoreCase);
      }
      catch
      {
        return false;
      }
    }

    private static LightingFixtureScheduleSyncPayload BuildLightingFixtureScheduleSyncPayload(
      LightingFixtureScheduleSyncSchedule schedule,
      string dwgPath,
      string tableHandle,
      string sourceApp
    )
    {
      LightingFixtureScheduleSyncSchedule normalizedSchedule = NormalizeSyncSchedule(schedule);
      string normalizedDwgPath = Path.GetFullPath(dwgPath ?? string.Empty);
      string projectBasePath = GetProjectBasePath(normalizedDwgPath);
      string projectId = ExtractProjectIdFromPath(normalizedDwgPath);
      string generatedAtUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);

      return new LightingFixtureScheduleSyncPayload
      {
        SchemaVersion = "1.0.0",
        Metadata = new LightingFixtureScheduleSyncMetadata
        {
          SourceApp = sourceApp ?? "autocad",
          GeneratedAtUtc = generatedAtUtc,
          GeneratedBy = Environment.UserName ?? "autocad",
          Fingerprint = ComputeLightingScheduleFingerprint(normalizedSchedule),
        },
        Project = new LightingFixtureScheduleSyncProject
        {
          ProjectId = projectId,
          ProjectBasePath = projectBasePath,
          DwgPath = normalizedDwgPath,
          DwgName = Path.GetFileName(normalizedDwgPath),
        },
        Table = new LightingFixtureScheduleSyncTable
        {
          TableHandle = tableHandle ?? string.Empty,
        },
        Schedule = normalizedSchedule,
      };
    }

    private static LightingFixtureScheduleSyncPayload NormalizeSyncPayload(
      LightingFixtureScheduleSyncPayload payload
    )
    {
      if (payload == null)
      {
        throw new InvalidDataException("Sync payload is missing.");
      }

      payload.Metadata = payload.Metadata ?? new LightingFixtureScheduleSyncMetadata();
      payload.Project = payload.Project ?? new LightingFixtureScheduleSyncProject();
      payload.Table = payload.Table ?? new LightingFixtureScheduleSyncTable();
      payload.Schedule = NormalizeSyncSchedule(payload.Schedule);
      if (string.IsNullOrWhiteSpace(payload.SchemaVersion))
      {
        payload.SchemaVersion = "1.0.0";
      }
      if (string.IsNullOrWhiteSpace(payload.Metadata.SourceApp))
      {
        payload.Metadata.SourceApp = "unknown";
      }
      if (string.IsNullOrWhiteSpace(payload.Metadata.GeneratedAtUtc))
      {
        payload.Metadata.GeneratedAtUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
      }
      if (string.IsNullOrWhiteSpace(payload.Metadata.GeneratedBy))
      {
        payload.Metadata.GeneratedBy = "unknown";
      }
      if (string.IsNullOrWhiteSpace(payload.Metadata.Fingerprint))
      {
        payload.Metadata.Fingerprint = ComputeLightingScheduleFingerprint(payload.Schedule);
      }
      return payload;
    }

    private static LightingFixtureScheduleSyncSchedule NormalizeSyncSchedule(
      LightingFixtureScheduleSyncSchedule schedule
    )
    {
      var normalizedRows = (schedule?.Rows ?? new List<LightingFixtureScheduleSyncRow>())
        .Select(NormalizeSyncRow)
        .ToList();
      if (normalizedRows.Count == 0)
      {
        normalizedRows.Add(new LightingFixtureScheduleSyncRow());
      }
      return new LightingFixtureScheduleSyncSchedule
      {
        Rows = normalizedRows,
        GeneralNotes = NormalizePlainText(schedule?.GeneralNotes),
        Notes = NormalizePlainText(schedule?.Notes),
        IncludeSymbolColumn = schedule?.IncludeSymbolColumn ?? false,
      };
    }

    private static LightingFixtureScheduleSyncRow NormalizeSyncRow(
      LightingFixtureScheduleSyncRow row
    )
    {
      var normalized = new LightingFixtureScheduleSyncRow
      {
        Mark = NormalizePlainText(row?.Mark),
        Description = NormalizePlainText(row?.Description),
        Manufacturer = NormalizePlainText(row?.Manufacturer),
        ModelNumber = NormalizePlainText(row?.ModelNumber),
        Mounting = NormalizePlainText(row?.Mounting),
        Volts = NormalizePlainText(row?.Volts),
        Watts = NormalizePlainText(row?.Watts),
        Notes = NormalizePlainText(row?.Notes),
        SymbolAssetPath = NormalizePlainText(row?.SymbolAssetPath),
        SymbolAlt = NormalizePlainText(row?.SymbolAlt),
        StarterFixtureKey = NormalizePlainText(row?.StarterFixtureKey),
      };
      RestoreCaliforniaStarterSymbolMetadata(normalized);
      return normalized;
    }

    private static void RestoreCaliforniaStarterSymbolMetadata(
      LightingFixtureScheduleSyncRow row
    )
    {
      if (
        row == null ||
        !string.IsNullOrWhiteSpace(row.SymbolAssetPath) ||
        !string.IsNullOrWhiteSpace(row.StarterFixtureKey)
      )
      {
        return;
      }

      string expectedModelFragment;
      if (
        !CaliforniaStarterModelFragments.TryGetValue(row.Mark, out expectedModelFragment) ||
        (row.ModelNumber ?? string.Empty).IndexOf(
          expectedModelFragment,
          StringComparison.OrdinalIgnoreCase
        ) < 0
      )
      {
        return;
      }

      string fixtureNumber = row.Mark.Substring(1);
      row.StarterFixtureKey = $"ca-2025-res-l{fixtureNumber}";
      row.SymbolAssetPath = $"assets/lighting/ca-residential-l{fixtureNumber}.png";
      if (string.IsNullOrWhiteSpace(row.SymbolAlt))
      {
        row.SymbolAlt = string.IsNullOrWhiteSpace(row.Description)
          ? $"Fixture {row.Mark} symbol"
          : $"{row.Description} symbol";
      }
    }

    private static string ComputeLightingScheduleFingerprint(
      LightingFixtureScheduleSyncSchedule schedule
    )
    {
      LightingFixtureScheduleSyncSchedule normalized = NormalizeSyncSchedule(schedule);
      var canonical = new
      {
        rows = normalized.Rows.Select(row => new
        {
          mark = row.Mark ?? string.Empty,
          description = row.Description ?? string.Empty,
          manufacturer = row.Manufacturer ?? string.Empty,
          modelNumber = row.ModelNumber ?? string.Empty,
          mounting = row.Mounting ?? string.Empty,
          volts = row.Volts ?? string.Empty,
          watts = row.Watts ?? string.Empty,
          notes = row.Notes ?? string.Empty,
        }).ToList(),
        generalNotes = normalized.GeneralNotes ?? string.Empty,
        notes = normalized.Notes ?? string.Empty,
      };

      string serialized = StableSerializeForFingerprint(canonical);
      uint hash = 0x811C9DC5;
      foreach (char c in serialized)
      {
        hash ^= c;
        hash *= 0x01000193;
      }
      return $"fnv1a:{hash:x8}";
    }

    private static string StableSerializeForFingerprint(object value)
    {
      JToken token = value == null ? JValue.CreateNull() : JToken.FromObject(value);
      JToken sorted = SortTokenKeys(token);
      return JsonConvert.SerializeObject(sorted, Formatting.None);
    }

    private static JToken SortTokenKeys(JToken token)
    {
      if (token == null)
      {
        return JValue.CreateNull();
      }

      JObject obj = token as JObject;
      if (obj != null)
      {
        var sortedObj = new JObject();
        foreach (JProperty prop in obj.Properties().OrderBy(p => p.Name, StringComparer.Ordinal))
        {
          sortedObj.Add(prop.Name, SortTokenKeys(prop.Value));
        }
        return sortedObj;
      }

      JArray arr = token as JArray;
      if (arr != null)
      {
        var sortedArr = new JArray();
        foreach (JToken item in arr)
        {
          sortedArr.Add(SortTokenKeys(item));
        }
        return sortedArr;
      }

      return token.DeepClone();
    }

    private static LightingFixtureScheduleSyncSchedule ExtractLightingScheduleFromTable(
      Table table
    )
    {
      if (table == null)
      {
        throw new InvalidOperationException("Table is required.");
      }

      LightingFixtureScheduleTableLayout layout = ResolveLightingScheduleTableLayout(table);
      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(table, layout.HeaderRow);
      var rows = new List<LightingFixtureScheduleSyncRow>();
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        rows.Add(new LightingFixtureScheduleSyncRow
        {
          Mark = GetCellPlainText(table, row, 0),
          Description = GetCellPlainText(table, row, GetLightingFixtureScheduleTextColumn(1, symbolColumn)),
          Manufacturer = GetCellPlainText(table, row, GetLightingFixtureScheduleTextColumn(2, symbolColumn)),
          ModelNumber = GetCellPlainText(table, row, GetLightingFixtureScheduleTextColumn(3, symbolColumn)),
          Mounting = GetCellPlainText(table, row, GetLightingFixtureScheduleTextColumn(4, symbolColumn)),
          Volts = GetCellPlainText(table, row, GetLightingFixtureScheduleTextColumn(5, symbolColumn)),
          Watts = GetCellPlainText(table, row, GetLightingFixtureScheduleTextColumn(6, symbolColumn)),
          Notes = GetCellPlainText(table, row, GetLightingFixtureScheduleTextColumn(7, symbolColumn)),
        });
      }

      var schedule = new LightingFixtureScheduleSyncSchedule
      {
        Rows = rows,
        GeneralNotes = StripNotesLabel(
          GetCellPlainText(table, layout.GeneralNotesRow, 0),
          "GENERAL NOTES"
        ),
        Notes = StripNotesLabel(
          GetCellPlainText(table, layout.NotesRow, 0),
          "NOTES"
        ),
        IncludeSymbolColumn = symbolColumn >= 0,
      };

      return NormalizeSyncSchedule(schedule);
    }

    private static LightingFixtureScheduleVisualNormalizationResult ApplyLightingScheduleToTable(
      Table table,
      Database db,
      Transaction tr,
      LightingFixtureScheduleSyncSchedule schedule
    )
    {
      if (table == null)
      {
        throw new InvalidOperationException("Table is required.");
      }
      if (db == null)
      {
        throw new InvalidOperationException("Database is required.");
      }
      if (tr == null)
      {
        throw new InvalidOperationException("Transaction is required.");
      }

      LightingFixtureScheduleSyncSchedule normalized = NormalizeSyncSchedule(schedule);
      LightingFixtureScheduleTableLayout layout = ResolveLightingScheduleTableLayout(table);
      EnsureLightingFixtureScheduleSymbolColumn(
        table,
        layout,
        normalized.IncludeSymbolColumn
      );
      layout = ResolveLightingScheduleTableLayout(table);
      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(table, layout.HeaderRow);
      int targetRowCount = Math.Max(1, normalized.Rows.Count);
      int existingRowCount = Math.Max(0, layout.GeneralNotesRow - layout.FirstDataRow);

      if (targetRowCount > existingRowCount)
      {
        int rowsToInsert = targetRowCount - existingRowCount;
        double rowHeight = ResolveDataRowHeight(table, layout);
        table.InsertRows(layout.GeneralNotesRow, rowHeight, rowsToInsert);
      }
      else if (targetRowCount < existingRowCount)
      {
        int rowsToDelete = existingRowCount - targetRowCount;
        table.DeleteRows(layout.FirstDataRow + targetRowCount, rowsToDelete);
      }

      layout = ResolveLightingScheduleTableLayout(table);
      for (int i = 0; i < targetRowCount; i++)
      {
        LightingFixtureScheduleSyncRow row = i < normalized.Rows.Count
          ? NormalizeSyncRow(normalized.Rows[i])
          : new LightingFixtureScheduleSyncRow();
        int tableRow = layout.FirstDataRow + i;
        SetCellPlainText(table, tableRow, 0, row.Mark);
        SetCellPlainText(table, tableRow, GetLightingFixtureScheduleTextColumn(1, symbolColumn), row.Description);
        SetCellPlainText(table, tableRow, GetLightingFixtureScheduleTextColumn(2, symbolColumn), row.Manufacturer);
        SetCellPlainText(table, tableRow, GetLightingFixtureScheduleTextColumn(3, symbolColumn), row.ModelNumber);
        SetCellPlainText(table, tableRow, GetLightingFixtureScheduleTextColumn(4, symbolColumn), row.Mounting);
        SetCellPlainText(table, tableRow, GetLightingFixtureScheduleTextColumn(5, symbolColumn), row.Volts);
        SetCellPlainText(table, tableRow, GetLightingFixtureScheduleTextColumn(6, symbolColumn), row.Watts);
        SetCellPlainText(table, tableRow, GetLightingFixtureScheduleTextColumn(7, symbolColumn), row.Notes);
        if (normalized.IncludeSymbolColumn)
        {
          SetLightingFixtureScheduleSymbolCell(table, db, tr, tableRow, row);
        }
      }

      SetCellPlainText(
        table,
        layout.GeneralNotesRow,
        0,
        ComposeNotesCellText("GENERAL NOTES", normalized.GeneralNotes)
      );
      SetCellPlainText(
        table,
        layout.NotesRow,
        0,
        ComposeNotesCellText("NOTES", normalized.Notes)
      );

      LightingFixtureScheduleVisualNormalizationResult result =
        ApplyTemplateVisualNormalization(table, db, tr, layout);
      if (normalized.IncludeSymbolColumn)
      {
        ApplyLightingFixtureScheduleSymbolColumnVisuals(table, layout);
      }
      return result;
    }

    private static void EnsureLightingFixtureScheduleSymbolColumn(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      bool includeSymbolColumn
    )
    {
      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(
        table,
        layout.HeaderRow
      );

      if (includeSymbolColumn && symbolColumn != 1)
      {
        if (symbolColumn >= 0)
        {
          table.DeleteColumns(symbolColumn, 1);
        }
        else if (table.NumColumns != 8)
        {
          throw new InvalidOperationException(
            "The symbol column can only be added to an eight-column lighting schedule."
          );
        }
        double width = table.Columns[0].Width;
        table.InsertColumns(1, width > 0 ? width : 0.75, 1);
        SetCellPlainText(table, layout.HeaderRow, 1, "SYMBOL");
        MergeLightingFixtureScheduleFullWidthRow(table, layout.HeaderRow - 1);
        MergeLightingFixtureScheduleFullWidthRow(table, layout.GeneralNotesRow);
        MergeLightingFixtureScheduleFullWidthRow(table, layout.NotesRow);
      }
      else if (!includeSymbolColumn && symbolColumn >= 0)
      {
        table.DeleteColumns(symbolColumn, 1);
        MergeLightingFixtureScheduleFullWidthRow(table, layout.HeaderRow - 1);
        MergeLightingFixtureScheduleFullWidthRow(table, layout.GeneralNotesRow);
        MergeLightingFixtureScheduleFullWidthRow(table, layout.NotesRow);
      }
    }

    private static int ResolveLightingFixtureScheduleSymbolColumnIndex(
      Table table,
      int headerRow
    )
    {
      if (table == null || headerRow < 0 || headerRow >= table.NumRows)
      {
        return -1;
      }
      foreach (int column in new[] { 1, 8 })
      {
        if (
          column < table.NumColumns &&
          string.Equals(
            NormalizeHeaderToken(GetCellPlainText(table, headerRow, column)),
            "SYMBOL",
            StringComparison.OrdinalIgnoreCase
          )
        )
        {
          return column;
        }
      }
      return -1;
    }

    private static int GetLightingFixtureScheduleTextColumn(
      int baseColumn,
      int symbolColumn
    )
    {
      return symbolColumn == 1 && baseColumn >= 1 ? baseColumn + 1 : baseColumn;
    }

    private static void MergeLightingFixtureScheduleFullWidthRow(Table table, int row)
    {
      if (row < 0 || row >= table.NumRows)
      {
        return;
      }
      CellRange existingRange;
      if (table.IsMergedCell(row, 0, out existingRange))
      {
        table.UnmergeCells(existingRange);
      }
      table.MergeCells(CellRange.Create(table, row, 0, row, table.NumColumns - 1));
    }

    private static void SetLightingFixtureScheduleSymbolCell(
      Table table,
      Database db,
      Transaction tr,
      int row,
      LightingFixtureScheduleSyncRow fixture
    )
    {
      const int symbolColumn = 1;
      SetCellPlainText(table, row, symbolColumn, string.Empty);
      ObjectId blockId = GetOrCreateLightingFixtureSymbolBlock(db, tr, fixture);
      if (blockId.IsNull)
      {
        return;
      }
      EnsureCellHasContent(table, row, symbolColumn);
      table.SetBlockTableRecordId(row, symbolColumn, 0, blockId, true);
      table.SetIsAutoScale(row, symbolColumn, 0, true);
      table.SetAlignment(row, symbolColumn, CellAlignment.MiddleCenter);
    }

    private static ObjectId GetOrCreateLightingFixtureSymbolBlock(
      Database db,
      Transaction tr,
      LightingFixtureScheduleSyncRow fixture
    )
    {
      if (db == null || tr == null || fixture == null)
      {
        return ObjectId.Null;
      }
      using (System.Drawing.Image preview = LoadLightingFixtureSymbolPreview(
        fixture.SymbolAssetPath,
        fixture.StarterFixtureKey
      ))
      {
        if (preview == null)
        {
          return ObjectId.Null;
        }

        string identity = $"{fixture.SymbolAssetPath}|{fixture.StarterFixtureKey}";
        string blockName = $"LFS_SYMBOL_V2_{ComputeLightingFixtureSymbolHash(identity)}";
        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        if (blockTable == null)
        {
          return ObjectId.Null;
        }
        if (blockTable.Has(blockName))
        {
          return blockTable[blockName];
        }

        blockTable.UpgradeOpen();
        var definition = new BlockTableRecord { Name = blockName };
        ObjectId definitionId = blockTable.Add(definition);
        tr.AddNewlyCreatedDBObject(definition, true);

        System.Drawing.Rectangle inkBounds = GetLightingFixtureSymbolInkBounds(preview);
        const int traceLongSide = 112;
        double traceScale = (double)traceLongSide /
          Math.Max(inkBounds.Width, inkBounds.Height);
        int width = Math.Max(1, (int)Math.Round(inkBounds.Width * traceScale));
        int height = Math.Max(1, (int)Math.Round(inkBounds.Height * traceScale));
        using (var bitmap = new System.Drawing.Bitmap(
          width,
          height,
          System.Drawing.Imaging.PixelFormat.Format32bppArgb
        ))
        {
          using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(bitmap))
          {
            graphics.Clear(System.Drawing.Color.White);
            graphics.CompositingQuality =
              System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            graphics.InterpolationMode =
              System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            graphics.DrawImage(
              preview,
              new System.Drawing.Rectangle(0, 0, width, height),
              inkBounds,
              System.Drawing.GraphicsUnit.Pixel
            );
          }

          double pixelSize = 1.0 / Math.Max(width, height);
          double totalWidth = width * pixelSize;
          double totalHeight = height * pixelSize;
          for (int y = 0; y < height; y++)
          {
            int x = 0;
            while (x < width)
            {
              while (x < width && !IsLightingFixtureSymbolDarkPixel(bitmap.GetPixel(x, y)))
              {
                x++;
              }
              int startX = x;
              while (x < width && IsLightingFixtureSymbolDarkPixel(bitmap.GetPixel(x, y)))
              {
                x++;
              }
              if (startX >= x)
              {
                continue;
              }

              double left = startX * pixelSize - totalWidth / 2.0;
              double right = x * pixelSize - totalWidth / 2.0;
              double top = totalHeight / 2.0 - y * pixelSize;
              double bottom = top - pixelSize;
              var solid = new Solid(
                new Point3d(left, bottom, 0),
                new Point3d(right, bottom, 0),
                new Point3d(left, top, 0),
                new Point3d(right, top, 0)
              );
              solid.Color = Color.FromColorIndex(ColorMethod.ByBlock, 0);
              definition.AppendEntity(solid);
              tr.AddNewlyCreatedDBObject(solid, true);
            }
          }
        }
        return definitionId;
      }
    }

    private static System.Drawing.Rectangle GetLightingFixtureSymbolInkBounds(
      System.Drawing.Image image
    )
    {
      using (var bitmap = new System.Drawing.Bitmap(image))
      {
        int minX = bitmap.Width;
        int minY = bitmap.Height;
        int maxX = -1;
        int maxY = -1;
        for (int y = 0; y < bitmap.Height; y++)
        {
          for (int x = 0; x < bitmap.Width; x++)
          {
            if (!IsLightingFixtureSymbolDarkPixel(bitmap.GetPixel(x, y)))
            {
              continue;
            }
            minX = Math.Min(minX, x);
            minY = Math.Min(minY, y);
            maxX = Math.Max(maxX, x);
            maxY = Math.Max(maxY, y);
          }
        }

        if (maxX < minX || maxY < minY)
        {
          return new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
        }

        const int padding = 2;
        minX = Math.Max(0, minX - padding);
        minY = Math.Max(0, minY - padding);
        maxX = Math.Min(bitmap.Width - 1, maxX + padding);
        maxY = Math.Min(bitmap.Height - 1, maxY + padding);
        return System.Drawing.Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
      }
    }

    private static bool IsLightingFixtureSymbolDarkPixel(System.Drawing.Color color)
    {
      if (color.A < 32)
      {
        return false;
      }
      int luminance = (299 * color.R + 587 * color.G + 114 * color.B) / 1000;
      return luminance < 180;
    }

    private static string ComputeLightingFixtureSymbolHash(string value)
    {
      uint hash = 2166136261;
      foreach (char c in value ?? string.Empty)
      {
        hash ^= c;
        hash *= 16777619;
      }
      return hash.ToString("X8", CultureInfo.InvariantCulture);
    }

    private static void ApplyLightingFixtureScheduleSymbolColumnVisuals(
      Table table,
      LightingFixtureScheduleTableLayout layout
    )
    {
      const int symbolColumn = 1;
      var warnings = new List<LightingFixtureScheduleWarning>();
      CopyLightingFixtureScheduleCellVisualProfile(
        table,
        layout.HeaderRow,
        0,
        layout.HeaderRow,
        symbolColumn,
        warnings
      );
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        CopyLightingFixtureScheduleCellVisualProfile(
          table,
          row,
          0,
          row,
          symbolColumn,
          warnings
        );
      }
    }

    private static void CopyLightingFixtureScheduleCellVisualProfile(
      Table table,
      int sourceRow,
      int sourceCol,
      int targetRow,
      int targetCol,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      CopyGridProfile(table, sourceRow, sourceCol, targetRow, targetCol, warnings);

      string cellStyle = SafeGet(
        warnings,
        "visual.symbol.cellStyle.get",
        () => table.GetCellStyle(sourceRow, sourceCol),
        sourceRow,
        sourceCol,
        string.Empty
      );
      if (!string.IsNullOrWhiteSpace(cellStyle))
      {
        SafeApply(
          warnings,
          "visual.symbol.cellStyle.set",
          () => table.SetCellStyle(targetRow, targetCol, cellStyle),
          targetRow,
          targetCol
        );
      }

      CellAlignment alignment = SafeGet(
        warnings,
        "visual.symbol.alignment.get",
        () => table.Alignment(sourceRow, sourceCol),
        sourceRow,
        sourceCol,
        CellAlignment.MiddleCenter
      );
      SafeApply(
        warnings,
        "visual.symbol.alignment.set",
        () => table.SetAlignment(targetRow, targetCol, alignment),
        targetRow,
        targetCol
      );

      ObjectId textStyleId = SafeGet(
        warnings,
        "visual.symbol.textStyle.get",
        () => table.TextStyle(sourceRow, sourceCol),
        sourceRow,
        sourceCol,
        ObjectId.Null
      );
      if (!textStyleId.IsNull)
      {
        SafeApply(
          warnings,
          "visual.symbol.textStyle.set",
          () => table.SetTextStyle(targetRow, targetCol, textStyleId),
          targetRow,
          targetCol
        );
      }

      double textHeight = SafeGet(
        warnings,
        "visual.symbol.textHeight.get",
        () => table.TextHeight(sourceRow, sourceCol),
        sourceRow,
        sourceCol,
        0.0
      );
      if (textHeight > 0)
      {
        SafeApply(
          warnings,
          "visual.symbol.textHeight.set",
          () => table.SetTextHeight(targetRow, targetCol, textHeight),
          targetRow,
          targetCol
        );
      }

      Color contentColor = SafeGet(
        warnings,
        "visual.symbol.contentColor.get",
        () => table.ContentColor(sourceRow, sourceCol),
        sourceRow,
        sourceCol,
        null as Color
      );
      if (contentColor != null)
      {
        SafeApply(
          warnings,
          "visual.symbol.contentColor.set",
          () => table.SetContentColor(targetRow, targetCol, contentColor),
          targetRow,
          targetCol
        );
      }

      bool backgroundColorNone = SafeGet(
        warnings,
        "visual.symbol.backgroundNone.get",
        () => table.IsBackgroundColorNone(sourceRow, sourceCol),
        sourceRow,
        sourceCol,
        true
      );
      SafeApply(
        warnings,
        "visual.symbol.backgroundNone.set",
        () => table.SetBackgroundColorNone(targetRow, targetCol, backgroundColorNone),
        targetRow,
        targetCol
      );
      if (!backgroundColorNone)
      {
        Color backgroundColor = SafeGet(
          warnings,
          "visual.symbol.backgroundColor.get",
          () => table.BackgroundColor(sourceRow, sourceCol),
          sourceRow,
          sourceCol,
          null as Color
        );
        if (backgroundColor != null)
        {
          SafeApply(
            warnings,
            "visual.symbol.backgroundColor.set",
            () => table.SetBackgroundColor(targetRow, targetCol, backgroundColor),
            targetRow,
            targetCol
          );
        }
      }

      foreach (CellMargins margin in new[]
      {
        CellMargins.Top,
        CellMargins.Left,
        CellMargins.Bottom,
        CellMargins.Right,
      })
      {
        CellMargins targetMargin = margin;
        double marginValue = SafeGet(
          warnings,
          "visual.symbol.margin.get",
          () => table.GetMargin(sourceRow, sourceCol, targetMargin),
          sourceRow,
          sourceCol,
          0.0
        );
        SafeApply(
          warnings,
          "visual.symbol.margin.set",
          () => table.SetMargin(targetRow, targetCol, targetMargin, marginValue),
          targetRow,
          targetCol
        );
      }
    }

    private static LightingFixtureScheduleVisualNormalizationResult ApplyTemplateVisualNormalization(
      Table table,
      Database db,
      Transaction tr,
      LightingFixtureScheduleTableLayout layout
    )
    {
      var warnings = new List<LightingFixtureScheduleWarning>();
      string templateError;
      LightingFixtureScheduleTemplateVisualProfile templateProfile =
        BuildTemplateVisualProfile(warnings, out templateError);

      if (templateProfile == null)
      {
        return ApplyRuntimeFallbackVisualNormalization(table, db, tr, layout, warnings, templateError);
      }

      bool usedRuntimeStyleFallback = false;
      ObjectId uniformTextStyleId = ObjectId.Null;
      if (
        !TryResolveTextStyleId(
          db,
          tr,
          templateProfile.UniformDataTextStyleName,
          out uniformTextStyleId
        )
      )
      {
        usedRuntimeStyleFallback = true;
        if (!TryGetRuntimeManufacturerTextStyleId(table, layout, out uniformTextStyleId))
        {
          uniformTextStyleId = db.Textstyle;
        }

        string missingStyleName = string.IsNullOrWhiteSpace(templateProfile.UniformDataTextStyleName)
          ? "(empty)"
          : templateProfile.UniformDataTextStyleName;
        AddWarning(
          warnings,
          "visual.textStyle",
          $"Template data text style '{missingStyleName}' was unavailable. Runtime baseline was used."
        );
      }

      Color uniformColor = ToAcadColor(
        templateProfile.UniformDataColor,
        warnings,
        "visual.dataColor",
        null,
        null
      );
      if (uniformColor == null)
      {
        usedRuntimeStyleFallback = true;
        if (!TryGetRuntimeManufacturerContentColor(table, layout, out uniformColor))
        {
          uniformColor = Color.FromColorIndex(ColorMethod.ByAci, 4);
        }
      }

      int maxDataColumns = Math.Min(8, table.NumColumns);
      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(
        table,
        layout.HeaderRow
      );
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        for (int col = 0; col < maxDataColumns; col++)
        {
          int targetCol = GetLightingFixtureScheduleTextColumn(col, symbolColumn);
          TableAnalyzeCell donorCell;
          if (templateProfile.DataCellsByColumn.TryGetValue(col, out donorCell))
          {
            ApplyCellVisualFromTemplate(table, db, tr, row, targetCol, donorCell, warnings);
          }
          else
          {
            CopyGridProfile(table, layout.FirstDataRow, targetCol, row, targetCol, warnings);
          }

          string text = GetCellPlainText(table, row, targetCol);
          ApplyUniformDataTextVisual(table, row, targetCol, uniformTextStyleId, uniformColor, warnings);
          SetCellPlainText(table, row, targetCol, text);
        }
      }
      ApplyUniformDataTextHeight(
        table,
        layout,
        table.NumColumns,
        templateProfile.UniformDataTextHeight,
        warnings
      );

      ApplyRowVisualFromTemplate(
        table,
        db,
        tr,
        layout.GeneralNotesRow,
        templateProfile.GeneralNotesCellsByColumn,
        warnings
      );
      ApplyRowVisualFromTemplate(
        table,
        db,
        tr,
        layout.NotesRow,
        templateProfile.NotesCellsByColumn,
        warnings
      );
      ApplyUniformNotesTextVisual(
        table,
        layout,
        table.NumColumns,
        uniformTextStyleId,
        uniformColor,
        warnings
      );
      ApplyUniformDataRowHeight(table, layout, templateProfile.UniformDataRowHeight, warnings);
      ApplyUniformDataPlacement(
        table,
        layout,
        maxDataColumns,
        templateProfile.DataPlacementByColumn,
        warnings
      );
      ApplyUniformBorderColor(
        table,
        Color.FromColorIndex(ColorMethod.ByAci, 1),
        warnings
      );

      return new LightingFixtureScheduleVisualNormalizationResult
      {
        Mode = usedRuntimeStyleFallback ? "template-with-style-fallback" : "template",
        UsedFallback = usedRuntimeStyleFallback,
        WarningCount = warnings.Count,
        Details = usedRuntimeStyleFallback
          ? "Template borders applied; runtime style baseline used."
          : "Template borders and text profile applied.",
      };
    }

    private static LightingFixtureScheduleVisualNormalizationResult ApplyRuntimeFallbackVisualNormalization(
      Table table,
      Database db,
      Transaction tr,
      LightingFixtureScheduleTableLayout layout,
      List<LightingFixtureScheduleWarning> warnings,
      string templateError
    )
    {
      ObjectId runtimeTextStyleId;
      if (!TryGetRuntimeManufacturerTextStyleId(table, layout, out runtimeTextStyleId))
      {
        runtimeTextStyleId = db.Textstyle;
      }

      Color runtimeColor;
      if (!TryGetRuntimeManufacturerContentColor(table, layout, out runtimeColor))
      {
        runtimeColor = Color.FromColorIndex(ColorMethod.ByAci, 4);
      }
      Dictionary<int, LightingFixtureScheduleDataPlacementProfile> runtimePlacementByColumn =
        ResolveRuntimeDataPlacementProfiles(table, layout);

      double runtimeTextHeight = ResolveMinimumPositiveDataTextHeight(
        table,
        layout,
        ResolveRuntimeManufacturerTextHeight(table, layout, 0.09375)
      );
      int maxDataColumns = Math.Min(8, table.NumColumns);
      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(
        table,
        layout.HeaderRow
      );
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        for (int col = 0; col < maxDataColumns; col++)
        {
          int targetCol = GetLightingFixtureScheduleTextColumn(col, symbolColumn);
          string text = GetCellPlainText(table, row, targetCol);
          CopyGridProfile(table, layout.FirstDataRow, targetCol, row, targetCol, warnings);
          ApplyUniformDataTextVisual(table, row, targetCol, runtimeTextStyleId, runtimeColor, warnings);
          SetCellPlainText(table, row, targetCol, text);
        }
      }
      ApplyUniformDataTextHeight(table, layout, table.NumColumns, runtimeTextHeight, warnings);

      int[] noteRows = { layout.GeneralNotesRow, layout.NotesRow };
      foreach (int noteRow in noteRows)
      {
        for (int col = 0; col < table.NumColumns; col++)
        {
          string text = GetCellPlainText(table, noteRow, col);
          CopyGridProfile(table, noteRow, col, noteRow, col, warnings);
          ApplyUniformDataTextVisual(table, noteRow, col, runtimeTextStyleId, runtimeColor, warnings);
          SetCellPlainText(table, noteRow, col, text);
        }
      }
      ApplyUniformDataRowHeight(
        table,
        layout,
        ResolveMinimumPositiveDataRowHeight(table, layout, ResolveDataRowHeight(table, layout)),
        warnings
      );
      ApplyUniformDataPlacement(
        table,
        layout,
        maxDataColumns,
        runtimePlacementByColumn,
        warnings
      );
      ApplyUniformBorderColor(
        table,
        Color.FromColorIndex(ColorMethod.ByAci, 1),
        warnings
      );

      if (!string.IsNullOrWhiteSpace(templateError))
      {
        AddWarning(
          warnings,
          "visual.template",
          $"Template profile could not be loaded. Runtime fallback was applied. Detail: {templateError}"
        );
      }

      return new LightingFixtureScheduleVisualNormalizationResult
      {
        Mode = "runtime-fallback",
        UsedFallback = true,
        WarningCount = warnings.Count,
        Details = "Runtime table profile was used for border/style normalization.",
      };
    }

    private static LightingFixtureScheduleTemplateVisualProfile BuildTemplateVisualProfile(
      List<LightingFixtureScheduleWarning> warnings,
      out string error
    )
    {
      error = string.Empty;
      try
      {
        string resourceName;
        TableAnalyzeExport template = LoadTemplateFromEmbeddedResource(out resourceName);
        ValidateTemplate(template, warnings);
        Dictionary<Tuple<int, int>, TableAnalyzeCell> cellMap = BuildTemplateCellMap(template);
        LightingFixtureScheduleTableLayout templateLayout = ResolveLightingScheduleTemplateLayout(
          template,
          cellMap
        );

        var profile = new LightingFixtureScheduleTemplateVisualProfile();
        int maxColumns = Math.Min(8, template.Table?.NumColumns ?? 8);
        for (int col = 0; col < maxColumns; col++)
        {
          TableAnalyzeCell donor = SelectTemplateDataDonorCell(cellMap, templateLayout, col);
          if (donor != null)
          {
            profile.DataCellsByColumn[col] = donor;
          }
          profile.DataPlacementByColumn[col] = ResolveTemplateDataPlacementProfile(
            donor,
            warnings,
            col
          );

          TableAnalyzeCell generalNotesCell;
          if (cellMap.TryGetValue(Tuple.Create(templateLayout.GeneralNotesRow, col), out generalNotesCell))
          {
            profile.GeneralNotesCellsByColumn[col] = generalNotesCell;
          }

          TableAnalyzeCell notesCell;
          if (cellMap.TryGetValue(Tuple.Create(templateLayout.NotesRow, col), out notesCell))
          {
            profile.NotesCellsByColumn[col] = notesCell;
          }
        }

        TableAnalyzeCell manufacturerCell;
        profile.DataCellsByColumn.TryGetValue(2, out manufacturerCell);
        profile.UniformDataTextStyleName = ResolveTemplateDataTextStyleName(
          manufacturerCell,
          profile.DataCellsByColumn.Values
        );
        profile.UniformDataColor = ResolveTemplateDataColor(
          manufacturerCell,
          profile.DataCellsByColumn.Values
        ) ?? new TableAnalyzeColor
        {
          ColorMethod = "ByAci",
          ColorIndex = 4,
          IsByAci = true
        };
        profile.UniformDataRowHeight = ResolveTemplateUniformDataRowHeight(templateLayout, template);
        profile.UniformDataTextHeight = ResolveTemplateUniformDataTextHeight(
          manufacturerCell,
          profile.DataCellsByColumn.Values
        );

        if (string.IsNullOrWhiteSpace(profile.UniformDataTextStyleName))
        {
          profile.UniformDataTextStyleName = "ARIALNARROW-1-8";
        }

        return profile;
      }
      catch (System.Exception ex)
      {
        error = ex.Message;
        return null;
      }
    }

    private static LightingFixtureScheduleTableLayout ResolveLightingScheduleTemplateLayout(
      TableAnalyzeExport template,
      Dictionary<Tuple<int, int>, TableAnalyzeCell> cellMap
    )
    {
      int numRows = template?.Table?.NumRows ?? 0;
      int numCols = template?.Table?.NumColumns ?? 0;
      if (numRows < 4 || numCols < 8)
      {
        throw new InvalidOperationException("Template dimensions are invalid for lighting schedule sync.");
      }

      int headerRow = -1;
      for (int row = 0; row < numRows; row++)
      {
        if (IsLightingScheduleTemplateHeaderRow(cellMap, row, numCols))
        {
          headerRow = row;
          break;
        }
      }

      if (headerRow < 0)
      {
        throw new InvalidOperationException(
          "Template header row could not be resolved for visual normalization."
        );
      }

      int generalNotesRow = FindTemplateRowByPrefix(cellMap, headerRow + 1, numRows, "GENERAL NOTES");
      if (generalNotesRow < 0 && numRows >= headerRow + 3)
      {
        generalNotesRow = numRows - 2;
      }

      int notesRow = FindTemplateRowByPrefix(
        cellMap,
        Math.Max(generalNotesRow + 1, headerRow + 1),
        numRows,
        "NOTES"
      );
      if (notesRow < 0 && numRows >= headerRow + 2)
      {
        notesRow = numRows - 1;
      }

      int firstDataRow = headerRow + 1;
      if (generalNotesRow <= firstDataRow || notesRow <= generalNotesRow)
      {
        throw new InvalidOperationException("Template row layout is invalid for lighting schedule sync.");
      }

      return new LightingFixtureScheduleTableLayout
      {
        HeaderRow = headerRow,
        FirstDataRow = firstDataRow,
        GeneralNotesRow = generalNotesRow,
        NotesRow = notesRow,
      };
    }

    private static Dictionary<Tuple<int, int>, TableAnalyzeCell> BuildTemplateCellMap(
      TableAnalyzeExport template
    )
    {
      return (template?.Cells ?? new List<TableAnalyzeCell>())
        .Where(cell => cell != null)
        .GroupBy(cell => Tuple.Create(cell.Row, cell.Column))
        .ToDictionary(group => group.Key, group => group.First());
    }

    private static bool IsLightingScheduleTemplateHeaderRow(
      Dictionary<Tuple<int, int>, TableAnalyzeCell> cellMap,
      int row,
      int numColumns
    )
    {
      if (row < 0 || numColumns < 8)
      {
        return false;
      }

      for (int col = 0; col < LightingFixtureScheduleHeaders.Length; col++)
      {
        TableAnalyzeCell cell;
        string text = cellMap.TryGetValue(Tuple.Create(row, col), out cell)
          ? NormalizeHeaderToken(GetTemplateCellPlainText(cell))
          : string.Empty;
        if (!string.Equals(text, LightingFixtureScheduleHeaders[col], StringComparison.OrdinalIgnoreCase))
        {
          return false;
        }
      }

      return true;
    }

    private static int FindTemplateRowByPrefix(
      Dictionary<Tuple<int, int>, TableAnalyzeCell> cellMap,
      int startRow,
      int numRows,
      string prefix
    )
    {
      int start = Math.Max(0, startRow);
      for (int row = start; row < numRows; row++)
      {
        TableAnalyzeCell cell;
        string text = cellMap.TryGetValue(Tuple.Create(row, 0), out cell)
          ? NormalizeHeaderToken(GetTemplateCellPlainText(cell))
          : string.Empty;
        if (text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
          return row;
        }
      }

      return -1;
    }

    private static string GetTemplateCellPlainText(TableAnalyzeCell cell)
    {
      if (cell?.Contents == null || cell.Contents.Count == 0)
      {
        return string.Empty;
      }

      foreach (
        TableAnalyzeCellContent content in cell.Contents
          .Where(item => item != null)
          .OrderBy(item => item.ContentIndex)
      )
      {
        string text = NormalizePlainText(GetPreferredContentText(content));
        if (!string.IsNullOrWhiteSpace(text))
        {
          return text;
        }
      }

      return string.Empty;
    }

    private static TableAnalyzeCell SelectTemplateDataDonorCell(
      Dictionary<Tuple<int, int>, TableAnalyzeCell> cellMap,
      LightingFixtureScheduleTableLayout layout,
      int col
    )
    {
      TableAnalyzeCell donor;
      if (cellMap.TryGetValue(Tuple.Create(layout.FirstDataRow, col), out donor) && donor != null)
      {
        return donor;
      }

      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        if (cellMap.TryGetValue(Tuple.Create(row, col), out donor) && donor != null)
        {
          return donor;
        }
      }

      if (cellMap.TryGetValue(Tuple.Create(layout.FirstDataRow, 2), out donor) && donor != null)
      {
        return donor;
      }

      return null;
    }

    private static string ResolveTemplateDataTextStyleName(
      TableAnalyzeCell manufacturerCell,
      IEnumerable<TableAnalyzeCell> dataDonors
    )
    {
      string styleName = ResolveCellTextStyleName(manufacturerCell);
      if (!string.IsNullOrWhiteSpace(styleName))
      {
        return styleName;
      }

      foreach (TableAnalyzeCell donor in dataDonors ?? Enumerable.Empty<TableAnalyzeCell>())
      {
        styleName = ResolveCellTextStyleName(donor);
        if (!string.IsNullOrWhiteSpace(styleName))
        {
          return styleName;
        }
      }

      return string.Empty;
    }

    private static string ResolveCellTextStyleName(TableAnalyzeCell cell)
    {
      if (cell == null)
      {
        return string.Empty;
      }

      TableAnalyzeCellContent firstContent = cell.Contents?
        .Where(content => content != null)
        .OrderBy(content => content.ContentIndex)
        .FirstOrDefault();
      if (!string.IsNullOrWhiteSpace(firstContent?.TextStyle?.Name))
      {
        return firstContent.TextStyle.Name;
      }

      return cell.TextStyle?.Name ?? string.Empty;
    }

    private static TableAnalyzeColor ResolveTemplateDataColor(
      TableAnalyzeCell manufacturerCell,
      IEnumerable<TableAnalyzeCell> dataDonors
    )
    {
      TableAnalyzeColor color = ResolveCellContentColor(manufacturerCell);
      if (color != null)
      {
        return color;
      }

      foreach (TableAnalyzeCell donor in dataDonors ?? Enumerable.Empty<TableAnalyzeCell>())
      {
        color = ResolveCellContentColor(donor);
        if (color != null)
        {
          return color;
        }
      }

      return null;
    }

    private static double ResolveTemplateUniformDataTextHeight(
      TableAnalyzeCell manufacturerCell,
      IEnumerable<TableAnalyzeCell> dataDonors
    )
    {
      double min = ResolveCellTextHeight(manufacturerCell);
      foreach (TableAnalyzeCell donor in dataDonors ?? Enumerable.Empty<TableAnalyzeCell>())
      {
        double donorHeight = ResolveCellTextHeight(donor);
        if (donorHeight > 0 && (min <= 0 || donorHeight < min))
        {
          min = donorHeight;
        }
      }

      return min > 0 ? min : 0.09375;
    }

    private static LightingFixtureScheduleDataPlacementProfile ResolveTemplateDataPlacementProfile(
      TableAnalyzeCell donorCell,
      List<LightingFixtureScheduleWarning> warnings,
      int col
    )
    {
      LightingFixtureScheduleDataPlacementProfile profile = CreateDefaultDataPlacementProfile(col);
      if (donorCell == null)
      {
        return profile;
      }

      profile.Alignment = ParseEnum(
        donorCell.Alignment,
        profile.Alignment,
        warnings,
        "visual.dataPlacement.alignment",
        null,
        col
      );
      profile.TextRotation = ParseEnum(
        donorCell.TextRotation,
        profile.TextRotation,
        warnings,
        "visual.dataPlacement.textRotation",
        null,
        col
      );
      profile.ContentLayout = ParseEnum(
        donorCell.ContentLayout,
        profile.ContentLayout,
        warnings,
        "visual.dataPlacement.contentLayout",
        null,
        col
      );
      if (donorCell.TopMargin.HasValue)
      {
        profile.TopMargin = donorCell.TopMargin.Value;
      }
      if (donorCell.LeftMargin.HasValue)
      {
        profile.LeftMargin = donorCell.LeftMargin.Value;
      }
      if (donorCell.BottomMargin.HasValue)
      {
        profile.BottomMargin = donorCell.BottomMargin.Value;
      }
      if (donorCell.RightMargin.HasValue)
      {
        profile.RightMargin = donorCell.RightMargin.Value;
      }

      return profile;
    }

    private static double ResolveTemplateUniformDataRowHeight(
      LightingFixtureScheduleTableLayout templateLayout,
      TableAnalyzeExport template
    )
    {
      if (templateLayout == null)
      {
        return 0.35;
      }

      List<double> rowHeights = template?.Table?.RowHeights;
      if (rowHeights == null || rowHeights.Count == 0)
      {
        return 0.35;
      }

      int start = Math.Max(0, templateLayout.FirstDataRow);
      int endExclusive = Math.Min(templateLayout.GeneralNotesRow, rowHeights.Count);
      if (start >= endExclusive)
      {
        return 0.35;
      }

      double min = double.MaxValue;
      for (int row = start; row < endExclusive; row++)
      {
        double h = rowHeights[row];
        if (h > 0 && h < min)
        {
          min = h;
        }
      }

      if (min < double.MaxValue)
      {
        return min;
      }

      return 0.35;
    }

    private static TableAnalyzeColor ResolveCellContentColor(TableAnalyzeCell cell)
    {
      if (cell == null)
      {
        return null;
      }

      TableAnalyzeCellContent firstContent = cell.Contents?
        .Where(content => content != null)
        .OrderBy(content => content.ContentIndex)
        .FirstOrDefault();
      if (firstContent?.ContentColor != null)
      {
        return firstContent.ContentColor;
      }

      return cell.ContentColor;
    }

    private static double ResolveCellTextHeight(TableAnalyzeCell cell)
    {
      if (cell == null)
      {
        return 0;
      }

      double min = double.MaxValue;
      if (cell.TextHeight.HasValue && cell.TextHeight.Value > 0)
      {
        min = cell.TextHeight.Value;
      }

      foreach (TableAnalyzeCellContent content in cell.Contents ?? Enumerable.Empty<TableAnalyzeCellContent>())
      {
        if (
          content != null
          && content.TextHeight.HasValue
          && content.TextHeight.Value > 0
          && content.TextHeight.Value < min
        )
        {
          min = content.TextHeight.Value;
        }
      }

      return min < double.MaxValue ? min : 0;
    }

    private static void ApplyRowVisualFromTemplate(
      Table table,
      Database db,
      Transaction tr,
      int row,
      Dictionary<int, TableAnalyzeCell> donorCellsByColumn,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      int maxColumns = Math.Min(8, table.NumColumns);
      for (int col = 0; col < maxColumns; col++)
      {
        TableAnalyzeCell donorCell;
        if (donorCellsByColumn.TryGetValue(col, out donorCell))
        {
          ApplyCellVisualFromTemplate(table, db, tr, row, col, donorCell, warnings);
        }
      }
    }

    private static void ApplyCellVisualFromTemplate(
      Table table,
      Database db,
      Transaction tr,
      int row,
      int col,
      TableAnalyzeCell donorCell,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (
        donorCell == null
        || row < 0
        || row >= table.NumRows
        || col < 0
        || col >= table.NumColumns
      )
      {
        return;
      }

      string text = GetCellPlainText(table, row, col);
      EnsureCellHasContent(table, row, col);
      ApplyCellProperties(table, db, tr, row, col, donorCell, warnings);
      ApplyGridLines(table, db, tr, row, col, donorCell, warnings);
      ApplyCellContentVisualOnlyFromTemplate(table, db, tr, row, col, donorCell, warnings);
      SetCellPlainText(table, row, col, text);
    }

    private static void ApplyCellContentVisualOnlyFromTemplate(
      Table table,
      Database db,
      Transaction tr,
      int row,
      int col,
      TableAnalyzeCell donorCell,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      List<TableAnalyzeCellContent> donorContents = donorCell?.Contents?
        .Where(content => content != null && content.ContentIndex >= 0)
        .OrderBy(content => content.ContentIndex)
        .ToList();
      if (donorContents == null || donorContents.Count == 0)
      {
        return;
      }

      int currentCount = EnsureCellHasContent(table, row, col);
      int requiredCount = Math.Max(1, donorContents.Max(content => content.ContentIndex) + 1);
      for (int i = currentCount; i < requiredCount; i++)
      {
        int idx = i;
        SafeApply(warnings, "visual.content.create", () => table.CreateContent(row, col, idx), row, col);
      }

      foreach (TableAnalyzeCellContent donorContent in donorContents)
      {
        int contentIndex = donorContent.ContentIndex;
        if (donorContent.TextHeight.HasValue)
        {
          SafeApply(
            warnings,
            "visual.content.textHeight",
            () => table.SetTextHeight(row, col, contentIndex, donorContent.TextHeight.Value),
            row,
            col
          );
        }

        ObjectId contentTextStyleId;
        if (TryResolveTextStyleId(db, tr, donorContent.TextStyle?.Name, out contentTextStyleId))
        {
          SafeApply(
            warnings,
            "visual.content.textStyle",
            () => table.SetTextStyleId(row, col, contentIndex, contentTextStyleId),
            row,
            col
          );
        }

        Color contentColor = ToAcadColor(
          donorContent.ContentColor,
          warnings,
          "visual.content.color",
          row,
          col
        );
        if (contentColor != null)
        {
          SafeApply(
            warnings,
            "visual.content.color",
            () => table.SetContentColor(row, col, contentIndex, contentColor),
            row,
            col
          );
        }

        if (donorContent.IsAutoScale.HasValue)
        {
          SafeApply(
            warnings,
            "visual.content.autoScale",
            () => table.SetIsAutoScale(row, col, contentIndex, donorContent.IsAutoScale.Value),
            row,
            col
          );
        }

        if (donorContent.Rotation.HasValue)
        {
          SafeApply(
            warnings,
            "visual.content.rotation",
            () => table.SetRotation(row, col, contentIndex, donorContent.Rotation.Value),
            row,
            col
          );
        }

        if (donorContent.Scale.HasValue)
        {
          SafeApply(
            warnings,
            "visual.content.scale",
            () => table.SetScale(row, col, contentIndex, donorContent.Scale.Value),
            row,
            col
          );
        }
      }
    }

    private static void ApplyUniformDataTextVisual(
      Table table,
      int row,
      int col,
      ObjectId textStyleId,
      Color contentColor,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      int contentCount = Math.Max(1, EnsureCellHasContent(table, row, col));
      if (!textStyleId.IsNull)
      {
        SafeApply(warnings, "visual.data.textStyle", () => table.SetTextStyle(row, col, textStyleId), row, col);
        for (int i = 0; i < contentCount; i++)
        {
          int idx = i;
          SafeApply(
            warnings,
            "visual.data.textStyleId",
            () => table.SetTextStyleId(row, col, idx, textStyleId),
            row,
            col
          );
        }
      }

      if (contentColor != null)
      {
        SafeApply(
          warnings,
          "visual.data.contentColor",
          () => table.SetContentColor(row, col, contentColor),
          row,
          col
        );
        for (int i = 0; i < contentCount; i++)
        {
          int idx = i;
          SafeApply(
            warnings,
            "visual.data.contentColorByIndex",
            () => table.SetContentColor(row, col, idx, contentColor),
            row,
            col
          );
        }
      }
    }

    private static void ApplyUniformDataTextHeight(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      int maxDataColumns,
      double targetTextHeight,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (
        table == null
        || layout == null
        || maxDataColumns <= 0
        || targetTextHeight <= 0
      )
      {
        return;
      }

      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        for (int col = 0; col < maxDataColumns; col++)
        {
          int contentCount = Math.Max(1, EnsureCellHasContent(table, row, col));
          SafeApply(
            warnings,
            "visual.data.textHeight",
            () => table.SetTextHeight(row, col, targetTextHeight),
            row,
            col
          );
          for (int i = 0; i < contentCount; i++)
          {
            int idx = i;
            SafeApply(
              warnings,
              "visual.data.textHeightByIndex",
              () => table.SetTextHeight(row, col, idx, targetTextHeight),
              row,
              col
            );
          }
        }
      }
    }

    private static void ApplyUniformDataPlacement(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      int maxDataColumns,
      IDictionary<int, LightingFixtureScheduleDataPlacementProfile> placementByColumn,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (table == null || layout == null || maxDataColumns <= 0)
      {
        return;
      }

      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(
        table,
        layout.HeaderRow
      );
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        for (int col = 0; col < maxDataColumns; col++)
        {
          int targetCol = GetLightingFixtureScheduleTextColumn(col, symbolColumn);
          LightingFixtureScheduleDataPlacementProfile profile;
          if (
            placementByColumn == null
            || !placementByColumn.TryGetValue(col, out profile)
            || profile == null
          )
          {
            profile = CreateDefaultDataPlacementProfile(col);
          }

          ApplyDataPlacementProfile(table, row, targetCol, profile, warnings);
        }
        if (symbolColumn == 1)
        {
          LightingFixtureScheduleDataPlacementProfile symbolProfile;
          if (
            placementByColumn == null ||
            !placementByColumn.TryGetValue(0, out symbolProfile) ||
            symbolProfile == null
          )
          {
            symbolProfile = CreateDefaultDataPlacementProfile(0);
          }
          ApplyDataPlacementProfile(table, row, symbolColumn, symbolProfile, warnings);
        }
      }
    }

    private static void ApplyDataPlacementProfile(
      Table table,
      int row,
      int col,
      LightingFixtureScheduleDataPlacementProfile profile,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (
        table == null
        || profile == null
        || row < 0
        || row >= table.NumRows
        || col < 0
        || col >= table.NumColumns
      )
      {
        return;
      }

      SafeApply(
        warnings,
        "visual.data.alignment",
        () => table.SetAlignment(row, col, profile.Alignment),
        row,
        col
      );
      SafeApply(
        warnings,
        "visual.data.textRotation",
        () => table.SetTextRotation(row, col, profile.TextRotation),
        row,
        col
      );
      SafeApply(
        warnings,
        "visual.data.contentLayout",
        () => table.SetContentLayout(row, col, profile.ContentLayout),
        row,
        col
      );
      SafeApply(
        warnings,
        "visual.data.marginTop",
        () => table.SetMargin(row, col, CellMargins.Top, profile.TopMargin),
        row,
        col
      );
      SafeApply(
        warnings,
        "visual.data.marginLeft",
        () => table.SetMargin(row, col, CellMargins.Left, profile.LeftMargin),
        row,
        col
      );
      SafeApply(
        warnings,
        "visual.data.marginBottom",
        () => table.SetMargin(row, col, CellMargins.Bottom, profile.BottomMargin),
        row,
        col
      );
      SafeApply(
        warnings,
        "visual.data.marginRight",
        () => table.SetMargin(row, col, CellMargins.Right, profile.RightMargin),
        row,
        col
      );
    }

    private static void ApplyUniformNotesTextVisual(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      int maxDataColumns,
      ObjectId textStyleId,
      Color contentColor,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (table == null || layout == null || maxDataColumns <= 0)
      {
        return;
      }

      int[] noteRows = { layout.GeneralNotesRow, layout.NotesRow };
      foreach (int row in noteRows)
      {
        if (row < 0 || row >= table.NumRows)
        {
          continue;
        }

        for (int col = 0; col < maxDataColumns; col++)
        {
          string text = GetCellPlainText(table, row, col);
          ApplyUniformDataTextVisual(table, row, col, textStyleId, contentColor, warnings);
          SetCellPlainText(table, row, col, text);
        }
      }
    }

    private static void ApplyUniformDataRowHeight(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      double targetHeight,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (
        table == null
        || layout == null
        || targetHeight <= 0
      )
      {
        return;
      }

      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        int targetRow = row;
        SafeApply(
          warnings,
          "visual.row.height.set",
          () => table.SetRowHeight(targetRow, targetHeight),
          targetRow,
          null
        );
      }
    }

    private static void ApplyUniformBorderColor(
      Table table,
      Color borderColor,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (table == null || borderColor == null)
      {
        return;
      }

      for (int row = 0; row < table.NumRows; row++)
      {
        for (int col = 0; col < table.NumColumns; col++)
        {
          foreach (GridLineType gridType in EnumerateGridTypes())
          {
            int targetRow = row;
            int targetCol = col;
            SafeApply(
              warnings,
              "visual.grid.color.forceRed",
              () => table.SetGridColor(targetRow, targetCol, gridType, borderColor),
              targetRow,
              targetCol
            );
          }
        }
      }
    }

    private static double ResolveMinimumPositiveDataRowHeight(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      double fallback
    )
    {
      if (table == null || layout == null)
      {
        return fallback > 0 ? fallback : 0.35;
      }

      double min = double.MaxValue;
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        try
        {
          double h = table.RowHeight(row);
          if (h > 0 && h < min)
          {
            min = h;
          }
        }
        catch
        {
          // Best-effort scan only.
        }
      }

      if (min < double.MaxValue)
      {
        return min;
      }

      return fallback > 0 ? fallback : 0.35;
    }

    private static double ResolveMinimumPositiveDataTextHeight(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      double fallback
    )
    {
      if (table == null || layout == null)
      {
        return fallback > 0 ? fallback : 0.09375;
      }

      double min = double.MaxValue;
      int maxDataColumns = Math.Min(8, table.NumColumns);
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        for (int col = 0; col < maxDataColumns; col++)
        {
          double cellTextHeight = ResolveCellTextHeight(table, row, col);
          if (cellTextHeight > 0 && cellTextHeight < min)
          {
            min = cellTextHeight;
          }
        }
      }

      if (min < double.MaxValue)
      {
        return min;
      }

      return fallback > 0 ? fallback : 0.09375;
    }

    private static Dictionary<int, LightingFixtureScheduleDataPlacementProfile> ResolveRuntimeDataPlacementProfiles(
      Table table,
      LightingFixtureScheduleTableLayout layout
    )
    {
      int maxDataColumns = Math.Min(8, table?.NumColumns ?? 8);
      int symbolColumn = table == null || layout == null
        ? -1
        : ResolveLightingFixtureScheduleSymbolColumnIndex(table, layout.HeaderRow);
      var profiles = new Dictionary<int, LightingFixtureScheduleDataPlacementProfile>();
      for (int col = 0; col < maxDataColumns; col++)
      {
        LightingFixtureScheduleDataPlacementProfile profile = null;
        if (table != null && layout != null)
        {
          for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
          {
            int targetCol = GetLightingFixtureScheduleTextColumn(col, symbolColumn);
            profile = TryCaptureRuntimeDataPlacementProfile(table, row, targetCol);
            if (profile != null)
            {
              break;
            }
          }
        }

        profiles[col] = profile ?? CreateDefaultDataPlacementProfile(col);
      }

      return profiles;
    }

    private static LightingFixtureScheduleDataPlacementProfile TryCaptureRuntimeDataPlacementProfile(
      Table table,
      int row,
      int col
    )
    {
      if (
        table == null
        || row < 0
        || row >= table.NumRows
        || col < 0
        || col >= table.NumColumns
      )
      {
        return null;
      }

      LightingFixtureScheduleDataPlacementProfile profile = CreateDefaultDataPlacementProfile(col);
      bool captured = false;

      try
      {
        profile.Alignment = table.Alignment(row, col);
        captured = true;
      }
      catch
      {
        // Best-effort scan only.
      }

      try
      {
        profile.TextRotation = table.TextRotation(row, col);
        captured = true;
      }
      catch
      {
        // Best-effort scan only.
      }

      try
      {
        profile.ContentLayout = table.GetContentLayout(row, col);
        captured = true;
      }
      catch
      {
        // Best-effort scan only.
      }

      try
      {
        profile.TopMargin = table.GetMargin(row, col, CellMargins.Top);
        captured = true;
      }
      catch
      {
        // Best-effort scan only.
      }

      try
      {
        profile.LeftMargin = table.GetMargin(row, col, CellMargins.Left);
        captured = true;
      }
      catch
      {
        // Best-effort scan only.
      }

      try
      {
        profile.BottomMargin = table.GetMargin(row, col, CellMargins.Bottom);
        captured = true;
      }
      catch
      {
        // Best-effort scan only.
      }

      try
      {
        profile.RightMargin = table.GetMargin(row, col, CellMargins.Right);
        captured = true;
      }
      catch
      {
        // Best-effort scan only.
      }

      return captured ? profile : null;
    }

    private static LightingFixtureScheduleDataPlacementProfile CreateDefaultDataPlacementProfile(int col)
    {
      return new LightingFixtureScheduleDataPlacementProfile
      {
        Alignment = col == 1 ? CellAlignment.MiddleLeft : CellAlignment.MiddleCenter,
        ContentLayout = CellContentLayout.Flow,
        TextRotation = RotationAngle.Degrees000,
        TopMargin = 0.06,
        LeftMargin = 0.06,
        BottomMargin = 0.06,
        RightMargin = 0.06
      };
    }

    private static double ResolveRuntimeManufacturerTextHeight(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      double fallback
    )
    {
      if (table == null || layout == null)
      {
        return fallback > 0 ? fallback : 0.09375;
      }

      int[] preferredColumns = { 2, 0, 1, 3, 4, 5, 6, 7 };
      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(
        table,
        layout.HeaderRow
      );
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        foreach (int baseColumn in preferredColumns)
        {
          int col = GetLightingFixtureScheduleTextColumn(baseColumn, symbolColumn);
          if (col < 0 || col >= table.NumColumns)
          {
            continue;
          }

          double textHeight = ResolveCellTextHeight(table, row, col);
          if (textHeight > 0)
          {
            return textHeight;
          }
        }
      }

      return fallback > 0 ? fallback : 0.09375;
    }

    private static double ResolveCellTextHeight(Table table, int row, int col)
    {
      if (
        table == null
        || row < 0
        || row >= table.NumRows
        || col < 0
        || col >= table.NumColumns
      )
      {
        return 0;
      }

      double min = double.MaxValue;
      try
      {
        double textHeight = table.TextHeight(row, col);
        if (textHeight > 0)
        {
          min = textHeight;
        }
      }
      catch
      {
        // Best-effort scan only.
      }

      int contentCount = 0;
      try
      {
        contentCount = table.GetNumberOfContents(row, col);
      }
      catch
      {
        contentCount = 0;
      }

      for (int i = 0; i < contentCount; i++)
      {
        try
        {
          double textHeight = table.GetTextHeight(row, col, i);
          if (textHeight > 0 && textHeight < min)
          {
            min = textHeight;
          }
        }
        catch
        {
          // Best-effort scan only.
        }
      }

      return min < double.MaxValue ? min : 0;
    }

    private static int EnsureCellHasContent(Table table, int row, int col)
    {
      if (
        table == null
        || row < 0
        || row >= table.NumRows
        || col < 0
        || col >= table.NumColumns
      )
      {
        return 0;
      }

      int count = 0;
      try
      {
        count = table.GetNumberOfContents(row, col);
      }
      catch
      {
        count = 0;
      }

      if (count > 0)
      {
        return count;
      }

      try
      {
        table.CreateContent(row, col, 0);
        return 1;
      }
      catch
      {
        return 0;
      }
    }

    private static bool TryResolveTextStyleId(
      Database db,
      Transaction tr,
      string styleName,
      out ObjectId textStyleId
    )
    {
      textStyleId = ObjectId.Null;
      if (db == null || tr == null || string.IsNullOrWhiteSpace(styleName))
      {
        return false;
      }

      TextStyleTable textStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
      if (textStyleTable == null)
      {
        return false;
      }

      if (textStyleTable.Has(styleName))
      {
        textStyleId = textStyleTable[styleName];
        return !textStyleId.IsNull;
      }

      foreach (ObjectId id in textStyleTable)
      {
        TextStyleTableRecord record = tr.GetObject(id, OpenMode.ForRead) as TextStyleTableRecord;
        if (record != null && string.Equals(record.Name, styleName, StringComparison.OrdinalIgnoreCase))
        {
          textStyleId = id;
          return !textStyleId.IsNull;
        }
      }

      return false;
    }

    private static bool TryGetRuntimeManufacturerTextStyleId(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      out ObjectId textStyleId
    )
    {
      textStyleId = ObjectId.Null;
      if (table == null || layout == null)
      {
        return false;
      }

      int row = layout.FirstDataRow;
      int col = GetLightingFixtureScheduleTextColumn(
        2,
        ResolveLightingFixtureScheduleSymbolColumnIndex(table, layout.HeaderRow)
      );
      try
      {
        if (table.GetNumberOfContents(row, col) > 0)
        {
          ObjectId id = table.GetTextStyleId(row, col, 0);
          if (!id.IsNull)
          {
            textStyleId = id;
            return true;
          }
        }
      }
      catch
      {
        // Best-effort fallback below.
      }

      try
      {
        Cell manufacturerCell = table.Cells[row, col];
        ObjectId? candidateStyleId = manufacturerCell?.TextStyleId;
        if (candidateStyleId.HasValue && !candidateStyleId.Value.IsNull)
        {
          textStyleId = candidateStyleId.Value;
          return true;
        }
      }
      catch
      {
        // Best-effort fallback only.
      }

      return false;
    }

    private static bool TryGetRuntimeManufacturerContentColor(
      Table table,
      LightingFixtureScheduleTableLayout layout,
      out Color contentColor
    )
    {
      contentColor = null;
      if (table == null || layout == null)
      {
        return false;
      }

      int row = layout.FirstDataRow;
      int col = GetLightingFixtureScheduleTextColumn(
        2,
        ResolveLightingFixtureScheduleSymbolColumnIndex(table, layout.HeaderRow)
      );
      try
      {
        if (table.GetNumberOfContents(row, col) > 0)
        {
          Color color = table.GetContentColor(row, col, 0);
          if (color != null)
          {
            contentColor = color;
            return true;
          }
        }
      }
      catch
      {
        // Best-effort fallback only.
      }

      return false;
    }

    private static void CopyGridProfile(
      Table table,
      int sourceRow,
      int sourceCol,
      int targetRow,
      int targetCol,
      List<LightingFixtureScheduleWarning> warnings
    )
    {
      if (
        table == null
        || sourceRow < 0
        || sourceRow >= table.NumRows
        || sourceCol < 0
        || sourceCol >= table.NumColumns
        || targetRow < 0
        || targetRow >= table.NumRows
        || targetCol < 0
        || targetCol >= table.NumColumns
      )
      {
        return;
      }

      foreach (GridLineType gridType in EnumerateGridTypes())
      {
        Visibility visibility = SafeGet(
          warnings,
          "visual.grid.visibility.get",
          () => table.GetGridVisibility(sourceRow, sourceCol, gridType),
          sourceRow,
          sourceCol,
          Visibility.Visible
        );
        GridLineStyle lineStyle = SafeGet(
          warnings,
          "visual.grid.lineStyle.get",
          () => table.GetGridLineStyle(sourceRow, sourceCol, gridType),
          sourceRow,
          sourceCol,
          GridLineStyle.Single
        );
        LineWeight lineWeight = SafeGet(
          warnings,
          "visual.grid.lineWeight.get",
          () => table.GetGridLineWeight(sourceRow, sourceCol, gridType),
          sourceRow,
          sourceCol,
          LineWeight.ByBlock
        );
        ObjectId linetypeId = SafeGet(
          warnings,
          "visual.grid.linetype.get",
          () => table.GetGridLinetype(sourceRow, sourceCol, gridType),
          sourceRow,
          sourceCol,
          ObjectId.Null
        );
        Color gridColor = SafeGet(
          warnings,
          "visual.grid.color.get",
          () => table.GetGridColor(sourceRow, sourceCol, gridType),
          sourceRow,
          sourceCol,
          null as Color
        );
        double? spacing = SafeGet<double?>(
          warnings,
          "visual.grid.spacing.get",
          () => table.GetGridDoubleLineSpacing(sourceRow, sourceCol, gridType),
          sourceRow,
          sourceCol,
          null
        );

        SafeApply(
          warnings,
          "visual.grid.visibility.set",
          () => table.SetGridVisibility(targetRow, targetCol, gridType, visibility),
          targetRow,
          targetCol
        );
        SafeApply(
          warnings,
          "visual.grid.lineStyle.set",
          () => table.SetGridLineStyle(targetRow, targetCol, gridType, lineStyle),
          targetRow,
          targetCol
        );
        SafeApply(
          warnings,
          "visual.grid.lineWeight.set",
          () => table.SetGridLineWeight(targetRow, targetCol, gridType, lineWeight),
          targetRow,
          targetCol
        );
        if (!linetypeId.IsNull)
        {
          SafeApply(
            warnings,
            "visual.grid.linetype.set",
            () => table.SetGridLinetype(targetRow, targetCol, gridType, linetypeId),
            targetRow,
            targetCol
          );
        }
        if (gridColor != null)
        {
          SafeApply(
            warnings,
            "visual.grid.color.set",
            () => table.SetGridColor(targetRow, targetCol, gridType, gridColor),
            targetRow,
            targetCol
          );
        }
        if (spacing.HasValue)
        {
          SafeApply(
            warnings,
            "visual.grid.spacing.set",
            () => table.SetGridDoubleLineSpacing(targetRow, targetCol, gridType, spacing.Value),
            targetRow,
            targetCol
          );
        }
      }
    }

    private static IEnumerable<GridLineType> EnumerateGridTypes()
    {
      foreach (GridLineType gridType in Enum.GetValues(typeof(GridLineType)))
      {
        if (gridType == GridLineType.InvalidGridLine)
        {
          continue;
        }

        yield return gridType;
      }
    }

    private static void TryRecomputeTableBlock(Table table)
    {
      if (table == null)
      {
        return;
      }

      try
      {
        MethodInfo recomputeWithArg = table
          .GetType()
          .GetMethod(
            "RecomputeTableBlock",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            null,
            new[] { typeof(bool) },
            null
          );
        if (recomputeWithArg != null)
        {
          recomputeWithArg.Invoke(table, new object[] { true });
          return;
        }

        MethodInfo recomputeNoArg = table
          .GetType()
          .GetMethod(
            "RecomputeTableBlock",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            null,
            Type.EmptyTypes,
            null
          );
        if (recomputeNoArg != null)
        {
          recomputeNoArg.Invoke(table, null);
        }
      }
      catch
      {
        // Optional refresh only.
      }
    }

    private static void SafeEditorRegen(Editor ed)
    {
      if (ed == null)
      {
        return;
      }

      try
      {
        ed.Regen();
      }
      catch
      {
        // Best-effort visual refresh only.
      }
    }

    private static string FormatNormalizationSummary(
      LightingFixtureScheduleVisualNormalizationResult result
    )
    {
      if (result == null)
      {
        return "none";
      }

      string mode = string.IsNullOrWhiteSpace(result.Mode) ? "unknown" : result.Mode;
      string details = string.IsNullOrWhiteSpace(result.Details) ? string.Empty : $" ({result.Details})";
      if (result.WarningCount > 0)
      {
        return $"{mode}{details}, warnings: {result.WarningCount}";
      }

      return $"{mode}{details}";
    }

    private static LightingFixtureScheduleTableLayout ResolveLightingScheduleTableLayout(
      Table table
    )
    {
      int numRows = table.NumRows;
      int numCols = table.NumColumns;
      if (numRows < 4 || numCols < 8)
      {
        throw new InvalidOperationException(
          "Selected table does not match lighting fixture schedule shape."
        );
      }

      int headerRow = -1;
      for (int row = 0; row < numRows; row++)
      {
        if (IsLightingScheduleHeaderRow(table, row))
        {
          headerRow = row;
          break;
        }
      }

      if (headerRow < 0)
      {
        throw new InvalidOperationException(
          "Could not find the lighting fixture schedule header row."
        );
      }

      int generalNotesRow = FindRowByPrefix(table, headerRow + 1, "GENERAL NOTES");
      if (generalNotesRow < 0 && numRows >= headerRow + 3)
      {
        generalNotesRow = numRows - 2;
      }

      int notesRow = FindRowByPrefix(table, Math.Max(generalNotesRow + 1, headerRow + 1), "NOTES");
      if (notesRow < 0 && numRows >= headerRow + 2)
      {
        notesRow = numRows - 1;
      }

      int firstDataRow = headerRow + 1;
      if (generalNotesRow <= firstDataRow || notesRow <= generalNotesRow)
      {
        throw new InvalidOperationException(
          "Lighting fixture schedule rows could not be resolved. Ensure the selected table is a valid schedule."
        );
      }

      return new LightingFixtureScheduleTableLayout
      {
        HeaderRow = headerRow,
        FirstDataRow = firstDataRow,
        GeneralNotesRow = generalNotesRow,
        NotesRow = notesRow,
      };
    }

    private static int FindRowByPrefix(Table table, int startRow, string prefix)
    {
      int start = Math.Max(0, startRow);
      for (int row = start; row < table.NumRows; row++)
      {
        string text = NormalizeHeaderToken(GetCellPlainText(table, row, 0));
        if (text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
          return row;
        }
      }
      return -1;
    }

    private static bool IsLightingScheduleHeaderRow(Table table, int row)
    {
      if (table == null || row < 0 || row >= table.NumRows || table.NumColumns < 8)
      {
        return false;
      }

      int symbolColumn = ResolveLightingFixtureScheduleSymbolColumnIndex(table, row);
      for (int col = 0; col < LightingFixtureScheduleHeaders.Length; col++)
      {
        int tableColumn = GetLightingFixtureScheduleTextColumn(col, symbolColumn);
        string text = NormalizeHeaderToken(GetCellPlainText(table, row, tableColumn));
        if (!string.Equals(text, LightingFixtureScheduleHeaders[col], StringComparison.OrdinalIgnoreCase))
        {
          return false;
        }
      }
      return true;
    }

    private static string NormalizeHeaderToken(string value)
    {
      string text = NormalizeHeaderTokenSourceText(value).ToUpperInvariant();
      if (text.EndsWith(":", StringComparison.Ordinal))
      {
        text = text.Substring(0, text.Length - 1).Trim();
      }
      return text.Trim();
    }

    private static string GetCellPlainText(Table table, int row, int col)
    {
      if (table == null || row < 0 || row >= table.NumRows || col < 0 || col >= table.NumColumns)
      {
        return string.Empty;
      }

      string plainText = string.Empty;
      string rawText = string.Empty;
      bool hasPlain = false;
      bool hasRaw = false;

      try
      {
        plainText = table.GetTextString(row, col, 0, FormatOption.IgnoreMtextFormat);
        hasPlain = true;
      }
      catch
      {
        // Best-effort fallback below.
      }

      try
      {
        // ForEditing tends to preserve logical paragraph markers more reliably than
        // the default formatter, which helps avoid Alt+Enter becoming double newlines.
        rawText = table.GetTextString(row, col, 0, FormatOption.ForEditing);
        hasRaw = true;
      }
      catch
      {
        // Best-effort fallback below.
      }

      if (!hasRaw)
      {
        try
        {
          rawText = table.GetTextString(row, col, 0);
          hasRaw = true;
        }
        catch
        {
          // Best-effort fallback below.
        }
      }

      if (!hasRaw)
      {
        string cellContentText = TryGetFirstCellContentText(table, row, col);
        if (!string.IsNullOrEmpty(cellContentText))
        {
          rawText = cellContentText;
          hasRaw = true;
        }
      }

      try
      {
        string text = table.Cells[row, col].TextString;
        if (!hasPlain)
        {
          plainText = text;
          hasPlain = true;
        }
      }
      catch
      {
        // Best-effort fallback below.
      }

      if (hasPlain || hasRaw)
      {
        return NormalizeTableCellText(plainText, rawText);
      }

      return string.Empty;
    }

    private static void SetCellPlainText(Table table, int row, int col, string text)
    {
      if (table == null || row < 0 || row >= table.NumRows || col < 0 || col >= table.NumColumns)
      {
        return;
      }

      string normalized = NormalizePlainText(text);
      try
      {
        int count = table.GetNumberOfContents(row, col);
        if (count <= 0)
        {
          table.CreateContent(row, col, 0);
          count = 1;
        }

        table.SetTextString(row, col, 0, normalized);
        for (int i = 1; i < count; i++)
        {
          table.SetTextString(row, col, i, string.Empty);
        }
      }
      catch
      {
        try
        {
          table.Cells[row, col].TextString = normalized;
        }
        catch
        {
          // Best-effort fallback only.
        }
      }
    }

    private static string TryGetFirstCellContentText(Table table, int row, int col)
    {
      if (table == null)
      {
        return string.Empty;
      }

      try
      {
        Cell cell = table.Cells[row, col];
        if (cell == null || cell.Contents == null || cell.Contents.Count <= 0)
        {
          return string.Empty;
        }

        CellContent content = cell.Contents[0];
        return content?.TextString ?? string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }

    private static double ResolveDataRowHeight(Table table, LightingFixtureScheduleTableLayout layout)
    {
      if (table == null)
      {
        return 0.35;
      }

      try
      {
        int candidate = layout.FirstDataRow;
        if (candidate >= 0 && candidate < table.NumRows)
        {
          double h = table.RowHeight(candidate);
          if (h > 0) return h;
        }
      }
      catch
      {
        // Best-effort fallback below.
      }

      try
      {
        int beforeNotes = layout.GeneralNotesRow - 1;
        if (beforeNotes >= 0 && beforeNotes < table.NumRows)
        {
          double h = table.RowHeight(beforeNotes);
          if (h > 0) return h;
        }
      }
      catch
      {
        // Best-effort fallback below.
      }

      return 0.35;
    }

    private static string StripNotesLabel(string value, string label)
    {
      string text = NormalizePlainText(value);
      if (string.IsNullOrEmpty(text))
      {
        return string.Empty;
      }
      string prefix = label ?? string.Empty;
      if (!string.IsNullOrEmpty(prefix) && text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
      {
        int idx = prefix.Length;
        if (idx < text.Length && text[idx] == ':')
        {
          idx++;
        }
        if (idx < text.Length && text[idx] == '\n')
        {
          idx++;
        }
        text = text.Substring(idx);
      }
      return text;
    }

    private static string ComposeNotesCellText(string label, string body)
    {
      string normalizedLabel = string.IsNullOrWhiteSpace(label) ? "NOTES" : label.Trim();
      string normalizedBody = NormalizePlainText(body);
      if (string.IsNullOrWhiteSpace(normalizedBody))
      {
        return $"{normalizedLabel}:";
      }
      return $"{normalizedLabel}:\n{normalizedBody}";
    }

    private static string NormalizePlainText(string value)
    {
      if (value == null)
      {
        return string.Empty;
      }
      return NormalizeLineEndingsForSync(DecodeMtextParagraphMarkers(value));
    }

    private static string NormalizeLineEndingsForSync(string value)
    {
      if (value == null)
      {
        return string.Empty;
      }

      return value
        .Replace("\r\n", "\n")
        .Replace("\r", "\n");
    }

    private static string DecodeMtextParagraphMarkers(string value)
    {
      if (string.IsNullOrEmpty(value))
      {
        return string.Empty;
      }

      var sb = new StringBuilder(value.Length);
      for (int i = 0; i < value.Length; i++)
      {
        char ch = value[i];
        if (ch == '\\' && i + 1 < value.Length)
        {
          char marker = value[i + 1];
          if (marker == 'P' || marker == 'p')
          {
            sb.Append('\n');
            i++;

            // Avoid duplicate line breaks when an explicit newline follows an MText paragraph marker.
            if (i + 1 < value.Length)
            {
              char next = value[i + 1];
              if (next == '\r')
              {
                i++;
                if (i + 1 < value.Length && value[i + 1] == '\n')
                {
                  i++;
                }
              }
              else if (next == '\n')
              {
                i++;
              }
            }
            continue;
          }

          if (marker == '~')
          {
            sb.Append(' ');
            i++;
            continue;
          }
        }

        sb.Append(ch);
      }

      return sb.ToString();
    }

    private static string NormalizeHeaderTokenSourceText(string value)
    {
      string text = NormalizeLineEndingsForSync(DecodeMtextParagraphMarkers(value ?? string.Empty));
      text = text
        .Replace("\n", " ")
        .Replace("\t", " ");
      while (text.Contains("  "))
      {
        text = text.Replace("  ", " ");
      }
      return text.Trim();
    }

    private static string NormalizeTableCellText(string plainTextCandidate, string rawTextCandidate)
    {
      string normalizedPlain = NormalizePlainText(plainTextCandidate);
      string normalizedRaw = NormalizePlainText(rawTextCandidate);
      string best = normalizedPlain.Length > 0 ? normalizedPlain : normalizedRaw;
      if (best.Length == 0)
      {
        return string.Empty;
      }

      if (!string.IsNullOrEmpty(rawTextCandidate))
      {
        int expectedBreaks = CountLogicalLineBreaksInRawText(rawTextCandidate);
        int actualBreaks = CountLineBreaks(best);
        if (actualBreaks > expectedBreaks)
        {
          best = ReduceExtraNewlineRuns(best, actualBreaks - expectedBreaks);
        }
      }

      return best;
    }

    private static int CountLogicalLineBreaksInRawText(string rawText)
    {
      if (string.IsNullOrEmpty(rawText))
      {
        return 0;
      }

      int breaks = 0;
      for (int i = 0; i < rawText.Length; i++)
      {
        char ch = rawText[i];
        if (ch == '\\' && i + 1 < rawText.Length)
        {
          char marker = rawText[i + 1];
          if (marker == 'P' || marker == 'p')
          {
            breaks++;
            i++;

            if (i + 1 < rawText.Length)
            {
              char next = rawText[i + 1];
              if (next == '\r')
              {
                i++;
                if (i + 1 < rawText.Length && rawText[i + 1] == '\n')
                {
                  i++;
                }
              }
              else if (next == '\n')
              {
                i++;
              }
            }
            continue;
          }

          if (marker == 'n' || marker == 'N')
          {
            breaks++;
            i++;
            continue;
          }

          if (marker == 'r' || marker == 'R')
          {
            breaks++;
            i++;

            if (
              i + 2 < rawText.Length
              && rawText[i + 1] == '\\'
              && (rawText[i + 2] == 'n' || rawText[i + 2] == 'N')
            )
            {
              i += 2;
            }
            continue;
          }
        }

        if (ch == '\r')
        {
          breaks++;
          if (i + 1 < rawText.Length && rawText[i + 1] == '\n')
          {
            i++;
          }
          continue;
        }

        if (ch == '\n')
        {
          breaks++;
        }
      }

      return breaks;
    }

    private static int CountLineBreaks(string text)
    {
      if (string.IsNullOrEmpty(text))
      {
        return 0;
      }

      int breaks = 0;
      for (int i = 0; i < text.Length; i++)
      {
        if (text[i] == '\n')
        {
          breaks++;
        }
      }
      return breaks;
    }

    private static string ReduceExtraNewlineRuns(string text, int removeCount)
    {
      if (string.IsNullOrEmpty(text) || removeCount <= 0)
      {
        return text ?? string.Empty;
      }

      var sb = new StringBuilder(text.Length);
      int i = 0;
      while (i < text.Length)
      {
        char ch = text[i];
        if (ch != '\n')
        {
          sb.Append(ch);
          i++;
          continue;
        }

        int runStart = i;
        while (i < text.Length && text[i] == '\n')
        {
          i++;
        }

        int runLength = i - runStart;
        int removable = runLength > 1 ? Math.Min(removeCount, runLength - 1) : 0;
        int keep = runLength - removable;
        removeCount -= removable;
        sb.Append('\n', keep);
      }

      return sb.ToString();
    }

    private static LightingFixtureScheduleSyncPayload ReadLightingFixtureScheduleSyncFile(string path)
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        throw new InvalidDataException("Sync file path is empty.");
      }
      if (!File.Exists(path))
      {
        throw new FileNotFoundException("Sync file was not found.", path);
      }

      try
      {
        string json = File.ReadAllText(path, Encoding.UTF8);
        LightingFixtureScheduleSyncPayload payload = JsonConvert.DeserializeObject<LightingFixtureScheduleSyncPayload>(json);
        if (payload == null)
        {
          throw new InvalidDataException("Sync file contains no payload.");
        }
        return payload;
      }
      catch (JsonException ex)
      {
        throw new InvalidDataException(
          $"Sync file JSON is invalid at '{path}': {ex.Message}"
        );
      }
    }

    private static void WriteLightingFixtureScheduleSyncFile(
      string path,
      LightingFixtureScheduleSyncPayload payload
    )
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        throw new InvalidDataException("Sync file path is empty.");
      }

      string directory = Path.GetDirectoryName(path);
      if (string.IsNullOrWhiteSpace(directory))
      {
        throw new InvalidDataException("Sync file directory could not be resolved.");
      }

      Directory.CreateDirectory(directory);
      LightingFixtureScheduleSyncPayload normalized = NormalizeSyncPayload(payload);
      JsonSerializerSettings settings = new JsonSerializerSettings
      {
        Formatting = Formatting.Indented,
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        NullValueHandling = NullValueHandling.Ignore,
      };

      string tempPath = Path.Combine(directory, $"{Guid.NewGuid():N}.tmp");
      try
      {
        string json = JsonConvert.SerializeObject(normalized, settings);
        File.WriteAllText(tempPath, json, Encoding.UTF8);
        if (File.Exists(path))
        {
          File.Delete(path);
        }
        File.Move(tempPath, path);
      }
      finally
      {
        if (File.Exists(tempPath))
        {
          File.Delete(tempPath);
        }
      }
    }

    private static string ExtractProjectIdFromPath(string rawPath)
    {
      if (string.IsNullOrWhiteSpace(rawPath))
      {
        return string.Empty;
      }

      string[] parts = rawPath
        .Replace("/", "\\")
        .Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
      for (int i = parts.Length - 1; i >= 0; i--)
      {
        Match match = LightingProjectSegmentRegex.Match(parts[i].Trim());
        if (match.Success)
        {
          return match.Groups[1].Value;
        }
      }
      return string.Empty;
    }

    private static string GetProjectBasePath(string rawPath)
    {
      if (string.IsNullOrWhiteSpace(rawPath))
      {
        return string.Empty;
      }

      string normalized = rawPath
        .Trim()
        .Trim('"')
        .Replace("/", "\\")
        .TrimEnd('\\');
      if (string.IsNullOrWhiteSpace(normalized))
      {
        return string.Empty;
      }

      string[] parts = normalized.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
      if (parts.Length == 0)
      {
        return string.Empty;
      }

      int baseIndex = -1;
      for (int i = parts.Length - 1; i >= 0; i--)
      {
        if (LightingProjectSegmentRegex.IsMatch(parts[i].Trim()))
        {
          baseIndex = i;
          break;
        }
      }

      int takeCount = parts.Length;
      if (baseIndex >= 0)
      {
        takeCount = baseIndex + 1;
      }
      else if (Path.HasExtension(parts[parts.Length - 1]))
      {
        takeCount = parts.Length - 1;
      }

      if (takeCount <= 0)
      {
        return string.Empty;
      }

      return string.Join("\\", parts.Take(takeCount));
    }
  }

  internal sealed class LightingFixtureScheduleTableLayout
  {
    public int HeaderRow { get; set; }
    public int FirstDataRow { get; set; }
    public int GeneralNotesRow { get; set; }
    public int NotesRow { get; set; }
  }

  internal sealed class LightingFixtureScheduleTemplateVisualProfile
  {
    public Dictionary<int, TableAnalyzeCell> DataCellsByColumn { get; } =
      new Dictionary<int, TableAnalyzeCell>();
    public Dictionary<int, LightingFixtureScheduleDataPlacementProfile> DataPlacementByColumn { get; } =
      new Dictionary<int, LightingFixtureScheduleDataPlacementProfile>();
    public Dictionary<int, TableAnalyzeCell> GeneralNotesCellsByColumn { get; } =
      new Dictionary<int, TableAnalyzeCell>();
    public Dictionary<int, TableAnalyzeCell> NotesCellsByColumn { get; } =
      new Dictionary<int, TableAnalyzeCell>();
    public string UniformDataTextStyleName { get; set; }
    public TableAnalyzeColor UniformDataColor { get; set; }
    public double UniformDataTextHeight { get; set; }
    public double UniformDataRowHeight { get; set; }
  }

  internal sealed class LightingFixtureScheduleDataPlacementProfile
  {
    public CellAlignment Alignment { get; set; }
    public RotationAngle TextRotation { get; set; }
    public CellContentLayout ContentLayout { get; set; }
    public double TopMargin { get; set; }
    public double LeftMargin { get; set; }
    public double BottomMargin { get; set; }
    public double RightMargin { get; set; }
  }

  internal sealed class LightingFixtureScheduleVisualNormalizationResult
  {
    public string Mode { get; set; }
    public bool UsedFallback { get; set; }
    public int WarningCount { get; set; }
    public string Details { get; set; }
  }

  public class LightingFixtureScheduleSyncPayload
  {
    public string SchemaVersion { get; set; }
    public LightingFixtureScheduleSyncMetadata Metadata { get; set; }
    public LightingFixtureScheduleSyncProject Project { get; set; }
    public LightingFixtureScheduleSyncTable Table { get; set; }
    public LightingFixtureScheduleSyncSchedule Schedule { get; set; }
  }

  public class LightingFixtureScheduleSyncMetadata
  {
    public string SourceApp { get; set; }
    public string GeneratedAtUtc { get; set; }
    public string GeneratedBy { get; set; }
    public string Fingerprint { get; set; }
  }

  public class LightingFixtureScheduleSyncProject
  {
    public string ProjectId { get; set; }
    public string ProjectBasePath { get; set; }
    public string DwgPath { get; set; }
    public string DwgName { get; set; }
  }

  public class LightingFixtureScheduleSyncTable
  {
    public string TableHandle { get; set; }
  }

  public class LightingFixtureScheduleSyncSchedule
  {
    public List<LightingFixtureScheduleSyncRow> Rows { get; set; }
    public string GeneralNotes { get; set; }
    public string Notes { get; set; }
    public bool IncludeSymbolColumn { get; set; }
  }

  public class LightingFixtureScheduleSyncRow
  {
    public string Mark { get; set; }
    public string Description { get; set; }
    public string Manufacturer { get; set; }
    public string ModelNumber { get; set; }
    public string Mounting { get; set; }
    public string Volts { get; set; }
    public string Watts { get; set; }
    public string Notes { get; set; }
    public string SymbolAssetPath { get; set; }
    public string SymbolAlt { get; set; }
    public string StarterFixtureKey { get; set; }
  }
}
