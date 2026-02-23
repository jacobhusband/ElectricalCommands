using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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

    [CommandMethod("LFSPULL", CommandFlags.Modal)]
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
        string syncPath = ResolveLightingFixtureScheduleSyncPath(doc, db);
        ObjectId tableId = PromptForLightingFixtureScheduleTable(
          ed,
          "\nSelect lighting fixture schedule table to export to sync file: "
        );
        if (tableId == ObjectId.Null)
        {
          ed.WriteMessage("\nLFSPULL cancelled.");
          return;
        }

        LightingFixtureScheduleSyncPayload payload;
        string tableHandle = "UNKNOWN";
        int rowCount = 0;

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

          payload = BuildLightingFixtureScheduleSyncPayload(
            schedule,
            dwgPath,
            tableHandle,
            "autocad"
          );

          tr.Commit();
        }

        WriteLightingFixtureScheduleSyncFile(syncPath, payload);
        ed.WriteMessage(
          $"\nLFSPULL complete. Handle: {tableHandle}, Rows: {rowCount}, Output: {syncPath}"
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

    [CommandMethod("LFSPUSH", CommandFlags.Modal)]
    public void LightingFixtureSchedulePush()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      try
      {
        RequireSavedDwgPath(doc, db);
        string syncPath = ResolveLightingFixtureScheduleSyncPath(doc, db);
        if (!File.Exists(syncPath))
        {
          ed.WriteMessage(
            $"\nLFSPUSH cancelled. Sync file was not found: {syncPath}. Run LFSPULL first or push from desktop."
          );
          return;
        }

        ObjectId tableId = PromptForLightingFixtureScheduleTable(
          ed,
          "\nSelect lighting fixture schedule table to update from sync file: "
        );
        if (tableId == ObjectId.Null)
        {
          ed.WriteMessage("\nLFSPUSH cancelled.");
          return;
        }

        LightingFixtureScheduleSyncPayload payload = ReadLightingFixtureScheduleSyncFile(syncPath);
        payload = NormalizeSyncPayload(payload);
        int rowCount = payload.Schedule?.Rows?.Count ?? 0;
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

          normalizationResult = ApplyLightingScheduleToTable(table, db, tr, payload.Schedule);
          table.GenerateLayout();
          TryRecomputeTableBlock(table);
          tableHandle = table.Handle.ToString();
          tr.Commit();
        }

        SafeEditorRegen(ed);
        ed.WriteMessage(
          $"\nLFSPUSH complete. Handle: {tableHandle}, Rows applied: {rowCount}, Input: {syncPath}. Visual normalization: {FormatNormalizationSummary(normalizationResult)}"
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
        string syncPath = ResolveLightingFixtureScheduleSyncPath(activeDoc, db);
        if (!File.Exists(syncPath))
        {
          return false;
        }

        LightingFixtureScheduleSyncPayload payload = NormalizeSyncPayload(
          ReadLightingFixtureScheduleSyncFile(syncPath)
        );
        Transaction tr = db.TransactionManager.TopTransaction;
        if (tr == null)
        {
          throw new InvalidOperationException(
            "An active transaction is required to apply lighting fixture schedule sync data."
          );
        }

        LightingFixtureScheduleVisualNormalizationResult normalizationResult =
          ApplyLightingScheduleToTable(table, db, tr, payload.Schedule);
        message =
          $"Applied sync data from {syncPath}. Visual normalization: {FormatNormalizationSummary(normalizationResult)}.";
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
      };
    }

    private static LightingFixtureScheduleSyncRow NormalizeSyncRow(
      LightingFixtureScheduleSyncRow row
    )
    {
      return new LightingFixtureScheduleSyncRow
      {
        Mark = NormalizePlainText(row?.Mark),
        Description = NormalizePlainText(row?.Description),
        Manufacturer = NormalizePlainText(row?.Manufacturer),
        ModelNumber = NormalizePlainText(row?.ModelNumber),
        Mounting = NormalizePlainText(row?.Mounting),
        Volts = NormalizePlainText(row?.Volts),
        Watts = NormalizePlainText(row?.Watts),
        Notes = NormalizePlainText(row?.Notes),
      };
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
      var rows = new List<LightingFixtureScheduleSyncRow>();
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        rows.Add(new LightingFixtureScheduleSyncRow
        {
          Mark = GetCellPlainText(table, row, 0),
          Description = GetCellPlainText(table, row, 1),
          Manufacturer = GetCellPlainText(table, row, 2),
          ModelNumber = GetCellPlainText(table, row, 3),
          Mounting = GetCellPlainText(table, row, 4),
          Volts = GetCellPlainText(table, row, 5),
          Watts = GetCellPlainText(table, row, 6),
          Notes = GetCellPlainText(table, row, 7),
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
        SetCellPlainText(table, tableRow, 1, row.Description);
        SetCellPlainText(table, tableRow, 2, row.Manufacturer);
        SetCellPlainText(table, tableRow, 3, row.ModelNumber);
        SetCellPlainText(table, tableRow, 4, row.Mounting);
        SetCellPlainText(table, tableRow, 5, row.Volts);
        SetCellPlainText(table, tableRow, 6, row.Watts);
        SetCellPlainText(table, tableRow, 7, row.Notes);
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

      return ApplyTemplateVisualNormalization(table, db, tr, layout);
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
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        for (int col = 0; col < maxDataColumns; col++)
        {
          TableAnalyzeCell donorCell;
          if (templateProfile.DataCellsByColumn.TryGetValue(col, out donorCell))
          {
            ApplyCellVisualFromTemplate(table, db, tr, row, col, donorCell, warnings);
          }
          else
          {
            CopyGridProfile(table, layout.FirstDataRow, col, row, col, warnings);
          }

          string text = GetCellPlainText(table, row, col);
          ApplyUniformDataTextVisual(table, row, col, uniformTextStyleId, uniformColor, warnings);
          SetCellPlainText(table, row, col, text);
        }
      }

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
        maxDataColumns,
        uniformTextStyleId,
        uniformColor,
        warnings
      );
      ApplyUniformDataRowHeight(table, layout, templateProfile.UniformDataRowHeight, warnings);
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

      int maxDataColumns = Math.Min(8, table.NumColumns);
      for (int row = layout.FirstDataRow; row < layout.GeneralNotesRow; row++)
      {
        for (int col = 0; col < maxDataColumns; col++)
        {
          string text = GetCellPlainText(table, row, col);
          CopyGridProfile(table, layout.FirstDataRow, col, row, col, warnings);
          ApplyUniformDataTextVisual(table, row, col, runtimeTextStyleId, runtimeColor, warnings);
          SetCellPlainText(table, row, col, text);
        }
      }

      int[] noteRows = { layout.GeneralNotesRow, layout.NotesRow };
      foreach (int noteRow in noteRows)
      {
        for (int col = 0; col < maxDataColumns; col++)
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
      int col = 2;
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
      int col = 2;
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

      for (int col = 0; col < LightingFixtureScheduleHeaders.Length; col++)
      {
        string text = NormalizeHeaderToken(GetCellPlainText(table, row, col));
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
    public Dictionary<int, TableAnalyzeCell> GeneralNotesCellsByColumn { get; } =
      new Dictionary<int, TableAnalyzeCell>();
    public Dictionary<int, TableAnalyzeCell> NotesCellsByColumn { get; } =
      new Dictionary<int, TableAnalyzeCell>();
    public string UniformDataTextStyleName { get; set; }
    public TableAnalyzeColor UniformDataColor { get; set; }
    public double UniformDataRowHeight { get; set; }
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
  }
}
