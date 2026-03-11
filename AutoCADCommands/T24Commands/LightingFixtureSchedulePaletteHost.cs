using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

[assembly: ExtensionApplication(typeof(ElectricalCommands.LightingFixtureSchedulePluginInitializer))]

namespace ElectricalCommands
{
  public class LightingFixtureSchedulePluginInitializer : IExtensionApplication
  {
    public void Initialize()
    {
      GeneralCommands.InitializeLightingFixtureScheduleAutoSync();
    }

    public void Terminate()
    {
      GeneralCommands.TerminateLightingFixtureScheduleAutoSync();
    }
  }

  public partial class GeneralCommands
  {
    private static readonly Guid LightingFixtureSchedulePaletteGuid =
      Guid.Parse("94F92D0E-2D7B-45E8-8611-B6B0F0EB222A");

    private static ILightingFixtureScheduleHost s_lightingFixtureScheduleHost;
    private static ILightingFixtureScheduleEditorControl s_lightingFixtureScheduleEditorControl;
    private static LightingFixtureScheduleBindingContext s_lightingFixtureSchedulePaletteContext;
    private static DateTime s_nextLightingFixtureScheduleSyncUtc = DateTime.MinValue;
    private static bool s_lightingFixtureScheduleAutoSyncInitialized;
    private static bool s_lightingFixtureScheduleAutoSyncBusy;

    internal static IReadOnlyList<string> GetLightingFixtureScheduleColumnHeaders()
    {
      return LightingFixtureScheduleHeaders;
    }

    internal static LightingFixtureScheduleSyncSchedule CreateDefaultLightingFixtureSchedule()
    {
      return NormalizeSyncSchedule(new LightingFixtureScheduleSyncSchedule());
    }

    internal static LightingFixtureScheduleSyncSchedule NormalizeLightingFixtureScheduleForExternalUse(
      LightingFixtureScheduleSyncSchedule schedule
    )
    {
      return NormalizeSyncSchedule(schedule);
    }

    internal static void InitializeLightingFixtureScheduleAutoSync()
    {
      if (s_lightingFixtureScheduleAutoSyncInitialized)
      {
        return;
      }

      Autodesk.AutoCAD.ApplicationServices.Core.Application.Idle +=
        LightingFixtureScheduleAutoSyncOnIdle;
      s_lightingFixtureScheduleAutoSyncInitialized = true;
    }

    internal static void TerminateLightingFixtureScheduleAutoSync()
    {
      if (!s_lightingFixtureScheduleAutoSyncInitialized)
      {
        return;
      }

      Autodesk.AutoCAD.ApplicationServices.Core.Application.Idle -=
        LightingFixtureScheduleAutoSyncOnIdle;
      s_lightingFixtureScheduleAutoSyncInitialized = false;

      ReleaseLightingFixtureScheduleHost();
      s_lightingFixtureSchedulePaletteContext = null;
    }

    public void LightingFixtureScheduleOpenPalette()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      string stage = "initialization";
      try
      {
        stage = "auto-sync initialization";
        InitializeLightingFixtureScheduleAutoSync();

        stage = "drawing path resolution";
        string dwgPath = RequireSavedDwgPath(doc, db);

        stage = "binding resolution";
        LightingFixtureScheduleBindingContext context =
          ResolveLightingFixtureScheduleBindingContext(doc, db, ed, dwgPath);

        stage = "central store load";
        LightingFixtureScheduleStoreRecord record =
          EnsureLightingFixtureScheduleStoreRecord(context);

        stage = "UI host creation";
        EnsureLightingFixtureScheduleHost(ed);

        if (
          s_lightingFixtureScheduleHost == null ||
          s_lightingFixtureScheduleHost.IsDisposed ||
          !IsLightingFixtureScheduleEditorControlReady()
        )
        {
          throw new InvalidOperationException("Lighting schedule UI host could not be created.");
        }

        s_lightingFixtureSchedulePaletteContext = context;

        stage = "record load";
        s_lightingFixtureScheduleEditorControl.LoadRecord(
          context,
          record,
          string.IsNullOrWhiteSpace(context.TableHandle)
            ? "Editor open. Link a table to drive the AutoCAD schedule."
            : "Editor open. Auto-saves write to the shared store and update the linked table."
        );

        stage = "host show";
        s_lightingFixtureScheduleHost.Show();
        ed.WriteMessage(
          $"\nLFS: opened lighting schedule using {s_lightingFixtureScheduleHost.HostKind}."
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage(
          $"\nLFS error during {stage}: {FormatLightingFixtureScheduleException(ex)}"
        );
      }
    }

    private static void EnsureLightingFixtureScheduleHost(Editor ed)
    {
      if (
        s_lightingFixtureScheduleHost != null &&
        !s_lightingFixtureScheduleHost.IsDisposed &&
        IsLightingFixtureScheduleEditorControlReady()
      )
      {
        return;
      }

      if (s_lightingFixtureScheduleHost != null && s_lightingFixtureScheduleHost.IsDisposed)
      {
        ReleaseLightingFixtureScheduleHost(keepControl: true);
      }

      EnsureLightingFixtureScheduleEditorControl(ed);
      if (!IsLightingFixtureScheduleEditorControlReady())
      {
        throw new InvalidOperationException("Lighting schedule editor control could not be created.");
      }

      System.Exception paletteFailure = null;
      try
      {
        DetachLightingFixtureScheduleControlFromParent();
        s_lightingFixtureScheduleHost = new LightingFixtureSchedulePaletteDockHost(
          s_lightingFixtureScheduleEditorControl.HostControl,
          LightingFixtureSchedulePaletteGuid
        );
        ed?.WriteMessage("\nLFS: using AutoCAD palette host.");
        return;
      }
      catch (System.Exception ex)
      {
        paletteFailure = ex;
        ed?.WriteMessage(
          $"\nLFS: palette host unavailable. {FormatLightingFixtureScheduleException(ex)}"
        );
        ReleaseLightingFixtureScheduleHost(keepControl: true);
        EnsureLightingFixtureScheduleEditorControl(ed);
      }

      System.Exception formFailure = null;
      try
      {
        DetachLightingFixtureScheduleControlFromParent();
        s_lightingFixtureScheduleHost = new LightingFixtureScheduleModelessFormHost(
          s_lightingFixtureScheduleEditorControl.HostControl
        );
        ed?.WriteMessage("\nLFS: using modeless WinForms fallback host.");
        return;
      }
      catch (System.Exception ex)
      {
        formFailure = ex;
        ReleaseLightingFixtureScheduleHost(keepControl: true);
      }

      throw new InvalidOperationException(
        "Unable to create any lighting schedule UI host. " +
        $"Palette failure: {FormatLightingFixtureScheduleException(paletteFailure)} " +
        $"Fallback failure: {FormatLightingFixtureScheduleException(formFailure)}"
      );
    }

    private static void EnsureLightingFixtureScheduleEditorControl(Editor ed)
    {
      if (IsLightingFixtureScheduleEditorControlReady())
      {
        return;
      }

      ReleaseLightingFixtureScheduleEditorControl();

      System.Exception richFailure = null;
      try
      {
        AttachLightingFixtureScheduleEditorControl(new LightingFixtureSchedulePaletteControl());
        ed?.WriteMessage("\nLFS: using primary schedule editor.");
        return;
      }
      catch (System.Exception ex)
      {
        richFailure = ex;
        ed?.WriteMessage(
          $"\nLFS: primary editor unavailable. {FormatLightingFixtureScheduleException(ex)}"
        );
        ReleaseLightingFixtureScheduleEditorControl();
      }

      System.Exception fallbackFailure = null;
      try
      {
        AttachLightingFixtureScheduleEditorControl(new LightingFixtureScheduleFallbackControl());
        ed?.WriteMessage("\nLFS: using simplified schedule editor.");
        return;
      }
      catch (System.Exception ex)
      {
        fallbackFailure = ex;
        ReleaseLightingFixtureScheduleEditorControl();
      }

      throw new InvalidOperationException(
        "Unable to create any lighting schedule editor. " +
        $"Primary editor failure: {FormatLightingFixtureScheduleException(richFailure)} " +
        $"Fallback editor failure: {FormatLightingFixtureScheduleException(fallbackFailure)}"
      );
    }

    private static void AttachLightingFixtureScheduleEditorControl(
      ILightingFixtureScheduleEditorControl editorControl
    )
    {
      if (editorControl == null)
      {
        throw new ArgumentNullException(nameof(editorControl));
      }

      s_lightingFixtureScheduleEditorControl = editorControl;
      s_lightingFixtureScheduleEditorControl.SaveRequested += LightingFixtureSchedulePaletteSaveRequested;
      s_lightingFixtureScheduleEditorControl.ReloadRequested += LightingFixtureSchedulePaletteReloadRequested;
      s_lightingFixtureScheduleEditorControl.LinkTableRequested += LightingFixtureSchedulePaletteLinkTableRequested;
      s_lightingFixtureScheduleEditorControl.CopyFromTableRequested += LightingFixtureSchedulePaletteCopyFromTableRequested;
      s_lightingFixtureScheduleEditorControl.PlaceTableRequested += LightingFixtureSchedulePalettePlaceTableRequested;
    }

    private static bool IsLightingFixtureScheduleEditorControlReady()
    {
      return
        s_lightingFixtureScheduleEditorControl != null &&
        s_lightingFixtureScheduleEditorControl.HostControl != null &&
        !s_lightingFixtureScheduleEditorControl.HostControl.IsDisposed;
    }

    private static void DetachLightingFixtureScheduleControlFromParent()
    {
      if (!IsLightingFixtureScheduleEditorControlReady())
      {
        return;
      }

      Control control = s_lightingFixtureScheduleEditorControl.HostControl;
      Control parent = control.Parent;
      if (parent != null)
      {
        parent.Controls.Remove(control);
      }
    }

    private static void ReleaseLightingFixtureScheduleEditorControl()
    {
      if (s_lightingFixtureScheduleEditorControl == null)
      {
        return;
      }

      if (s_lightingFixtureScheduleEditorControl.HostControl != null)
      {
        Control control = s_lightingFixtureScheduleEditorControl.HostControl;
        if (!control.IsDisposed)
        {
          control.Dispose();
        }
      }

      s_lightingFixtureScheduleEditorControl = null;
    }

    private static void ReleaseLightingFixtureScheduleHost(bool keepControl = false)
    {
      if (s_lightingFixtureScheduleHost != null)
      {
        try
        {
          s_lightingFixtureScheduleHost.Hide();
        }
        catch
        {
        }

        s_lightingFixtureScheduleHost.Dispose();
        s_lightingFixtureScheduleHost = null;
      }

      if (!keepControl)
      {
        ReleaseLightingFixtureScheduleEditorControl();
      }
    }

    private static string FormatLightingFixtureScheduleException(System.Exception ex)
    {
      if (ex == null)
      {
        return "Unknown error.";
      }

      string message = $"{ex.GetType().Name}: {ex.Message}";
      if (ex.InnerException != null)
      {
        message += $" | Inner: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}";
      }
      return message;
    }

    private static void SetLightingFixtureScheduleEditorStatus(string message)
    {
      if (IsLightingFixtureScheduleEditorControlReady())
      {
        s_lightingFixtureScheduleEditorControl.SetStatus(message);
      }
    }

    private static LightingFixtureScheduleBindingContext ResolveLightingFixtureScheduleBindingContext(
      Document doc,
      Database db,
      Editor ed,
      string dwgPath
    )
    {
      LightingFixtureScheduleStoreLink link =
        GetLightingFixtureScheduleStoreLinkByDwgPath(dwgPath);
      string projectId =
        NormalizePlainText(link?.ProjectId) ??
        string.Empty;
      if (string.IsNullOrWhiteSpace(projectId))
      {
        projectId = ResolveLightingFixtureScheduleProjectIdForDrawing(doc, db, dwgPath);
      }
      if (string.IsNullOrWhiteSpace(projectId))
      {
        projectId = PromptForLightingFixtureScheduleProjectId(ed);
      }
      if (string.IsNullOrWhiteSpace(projectId))
      {
        throw new InvalidOperationException(
          "A project ID is required before the lighting schedule can link to the shared store."
        );
      }

      string tableHandle = NormalizePlainText(link?.TableHandle);
      ObjectId tableId = ResolveLightingFixtureScheduleTableIdByHandle(db, tableHandle);
      if (tableId == ObjectId.Null)
      {
        tableHandle = string.Empty;
      }

      return new LightingFixtureScheduleBindingContext
      {
        Document = doc,
        Database = db,
        Editor = ed,
        ProjectId = projectId,
        DwgPath = dwgPath,
        TableId = tableId,
        TableHandle = tableHandle,
      };
    }

    private static LightingFixtureScheduleStoreRecord EnsureLightingFixtureScheduleStoreRecord(
      LightingFixtureScheduleBindingContext context
    )
    {
      LightingFixtureScheduleStoreRecord record =
        GetLightingFixtureScheduleStoreRecord(context.ProjectId)
        ?? GetLightingFixtureScheduleStoreRecordForDwg(context.DwgPath);

      if (record == null)
      {
        LightingFixtureScheduleSyncSchedule seed = CreateDefaultLightingFixtureSchedule();
        if (context.TableId != ObjectId.Null)
        {
          using (Transaction tr = context.Database.TransactionManager.StartTransaction())
          {
            Table table = tr.GetObject(context.TableId, OpenMode.ForRead, false) as Table;
            if (table != null)
            {
              seed = ExtractLightingScheduleFromTable(table);
              context.TableHandle = table.Handle.ToString();
            }
            tr.Commit();
          }
        }

        record = SaveLightingFixtureScheduleStoreRecord(
          context.ProjectId,
          seed,
          context.DwgPath,
          context.TableHandle,
          LightingFixtureScheduleStoreUpdatedByAutoCAD,
          null,
          0
        );
      }

      if (!string.IsNullOrWhiteSpace(record?.TableHandle) && string.IsNullOrWhiteSpace(context.TableHandle))
      {
        context.TableHandle = record.TableHandle;
        context.TableId = ResolveLightingFixtureScheduleTableIdByHandle(context.Database, context.TableHandle);
      }

      SaveLightingFixtureScheduleStoreLink(
        context.ProjectId,
        context.DwgPath,
        context.TableHandle,
        record?.Version ?? 0
      );

      return record;
    }

    private static string PromptForLightingFixtureScheduleProjectId(Editor ed)
    {
      if (ed == null)
      {
        return string.Empty;
      }

      PromptStringOptions options = new PromptStringOptions(
        "\nEnter project ID to bind the lighting fixture schedule: "
      );
      options.AllowSpaces = true;
      PromptResult result = ed.GetString(options);
      return result.Status == PromptStatus.OK ? NormalizePlainText(result.StringResult) : string.Empty;
    }

    private static void LightingFixtureSchedulePaletteSaveRequested(
      object sender,
      LightingFixtureScheduleSaveRequestedEventArgs e
    )
    {
      LightingFixtureScheduleBindingContext context = s_lightingFixtureSchedulePaletteContext;
      if (context == null)
      {
        return;
      }

      try
      {
        LightingFixtureScheduleStoreRecord savedRecord =
          SaveLightingFixtureScheduleStoreRecord(
            context.ProjectId,
            e.Schedule,
            context.DwgPath,
            context.TableHandle,
            LightingFixtureScheduleStoreUpdatedByAutoCAD,
            e.LoadedVersion,
            e.LoadedVersion
          );
        if (!string.IsNullOrWhiteSpace(context.TableHandle))
        {
          ApplyLightingFixtureScheduleRecordToLinkedTable(context, savedRecord);
        }
        if (IsLightingFixtureScheduleEditorControlReady())
        {
          s_lightingFixtureScheduleEditorControl.LoadRecord(
            context,
            savedRecord,
            savedRecord.HadConflict
              ? $"Saved version {savedRecord.Version}. A newer central version was overwritten."
              : $"Saved version {savedRecord.Version}."
          );
        }
      }
      catch (System.Exception ex)
      {
        SetLightingFixtureScheduleEditorStatus($"Save failed: {ex.Message}");
      }
    }

    private static void LightingFixtureSchedulePaletteReloadRequested(object sender, EventArgs e)
    {
      LightingFixtureScheduleBindingContext context = s_lightingFixtureSchedulePaletteContext;
      if (context == null)
      {
        return;
      }

      try
      {
        LightingFixtureScheduleStoreRecord record =
          GetLightingFixtureScheduleStoreRecord(context.ProjectId);
        if (record == null)
        {
          SetLightingFixtureScheduleEditorStatus("No central schedule record was found.");
          return;
        }

        if (!string.IsNullOrWhiteSpace(context.TableHandle))
        {
          ApplyLightingFixtureScheduleRecordToLinkedTable(context, record);
        }
        if (IsLightingFixtureScheduleEditorControlReady())
        {
          s_lightingFixtureScheduleEditorControl.LoadRecord(
            context,
            record,
            $"Reloaded central version {record.Version}."
          );
        }
      }
      catch (System.Exception ex)
      {
        SetLightingFixtureScheduleEditorStatus($"Reload failed: {ex.Message}");
      }
    }

    private static void LightingFixtureSchedulePaletteLinkTableRequested(object sender, EventArgs e)
    {
      LightingFixtureScheduleBindingContext context = s_lightingFixtureSchedulePaletteContext;
      if (context == null || context.Editor == null)
      {
        return;
      }

      try
      {
        ObjectId tableId = PromptForLightingFixtureScheduleTable(
          context.Editor,
          "\nSelect lighting fixture schedule table to link to the central store: "
        );
        if (tableId == ObjectId.Null)
        {
          SetLightingFixtureScheduleEditorStatus("Link cancelled.");
          return;
        }

        using (Transaction tr = context.Database.TransactionManager.StartTransaction())
        {
          Table table = tr.GetObject(tableId, OpenMode.ForRead, false) as Table;
          if (table == null)
          {
            throw new InvalidOperationException("Selected object is not an AutoCAD table.");
          }
          context.TableId = tableId;
          context.TableHandle = table.Handle.ToString();
          tr.Commit();
        }

        LightingFixtureScheduleStoreRecord record =
          GetLightingFixtureScheduleStoreRecord(context.ProjectId)
          ?? EnsureLightingFixtureScheduleStoreRecord(context);
        ApplyLightingFixtureScheduleRecordToLinkedTable(context, record);
        if (IsLightingFixtureScheduleEditorControlReady())
        {
          s_lightingFixtureScheduleEditorControl.LoadRecord(
            context,
            record,
            "Table linked. Central-store updates now refresh this table automatically."
          );
        }
      }
      catch (System.Exception ex)
      {
        SetLightingFixtureScheduleEditorStatus($"Link failed: {ex.Message}");
      }
    }

    private static void LightingFixtureSchedulePaletteCopyFromTableRequested(
      object sender,
      EventArgs e
    )
    {
      LightingFixtureScheduleBindingContext context = s_lightingFixtureSchedulePaletteContext;
      if (context == null || context.Editor == null)
      {
        return;
      }

      try
      {
        ObjectId tableId = PromptForLightingFixtureScheduleTable(
          context.Editor,
          "\nSelect lighting fixture schedule table to copy into the editor: "
        );
        if (tableId == ObjectId.Null)
        {
          SetLightingFixtureScheduleEditorStatus("Copy cancelled.");
          return;
        }

        LightingFixtureScheduleStoreRecord record =
          CopyLightingFixtureScheduleFromTable(context, tableId);
        if (IsLightingFixtureScheduleEditorControlReady())
        {
          s_lightingFixtureScheduleEditorControl.LoadRecord(
            context,
            record,
            $"Copied and linked table {context.TableHandle}. Version {record?.Version ?? 0}."
          );
        }
      }
      catch (System.Exception ex)
      {
        SetLightingFixtureScheduleEditorStatus($"Copy failed: {ex.Message}");
      }
    }

    private static void LightingFixtureSchedulePalettePlaceTableRequested(
      object sender,
      LightingFixtureScheduleSaveRequestedEventArgs e
    )
    {
      LightingFixtureScheduleBindingContext context = s_lightingFixtureSchedulePaletteContext;
      if (context == null || context.Editor == null)
      {
        return;
      }

      try
      {
        string statusMessage;
        LightingFixtureScheduleStoreRecord record = PlaceLightingFixtureScheduleTable(
          context,
          e?.Schedule,
          e?.LoadedVersion ?? 0,
          out statusMessage
        );

        if (record == null)
        {
          SetLightingFixtureScheduleEditorStatus(statusMessage);
          return;
        }

        if (IsLightingFixtureScheduleEditorControlReady())
        {
          s_lightingFixtureScheduleEditorControl.LoadRecord(
            context,
            record,
            statusMessage
          );
        }
      }
      catch (System.Exception ex)
      {
        SetLightingFixtureScheduleEditorStatus($"Place failed: {ex.Message}");
      }
    }

    private static LightingFixtureScheduleStoreRecord CopyLightingFixtureScheduleFromTable(
      LightingFixtureScheduleBindingContext context,
      ObjectId tableId
    )
    {
      if (context == null || context.Database == null)
      {
        throw new InvalidOperationException("Lighting schedule context is unavailable.");
      }

      string tableHandle = string.Empty;
      LightingFixtureScheduleSyncSchedule schedule = null;
      using (Transaction tr = context.Database.TransactionManager.StartTransaction())
      {
        Table table = tr.GetObject(tableId, OpenMode.ForRead, false) as Table;
        if (table == null)
        {
          throw new InvalidOperationException("Selected object is not an AutoCAD table.");
        }

        schedule = NormalizeSyncSchedule(ExtractLightingScheduleFromTable(table));
        tableHandle = table.Handle.ToString();
        tr.Commit();
      }

      LightingFixtureScheduleStoreRecord record = SaveLightingFixtureScheduleStoreRecord(
        context.ProjectId,
        schedule,
        context.DwgPath,
        tableHandle,
        LightingFixtureScheduleStoreUpdatedByAutoCAD,
        null,
        0
      );

      context.TableId = tableId;
      context.TableHandle = tableHandle;
      SaveLightingFixtureScheduleStoreLink(
        context.ProjectId,
        context.DwgPath,
        tableHandle,
        record?.Version ?? 0
      );

      return record;
    }

    private static LightingFixtureScheduleStoreRecord PlaceLightingFixtureScheduleTable(
      LightingFixtureScheduleBindingContext context,
      LightingFixtureScheduleSyncSchedule schedule,
      long loadedVersion,
      out string statusMessage
    )
    {
      if (context == null || context.Database == null || context.Editor == null)
      {
        throw new InvalidOperationException("Lighting schedule context is unavailable.");
      }

      string resourceName = string.Empty;
      var warnings = new List<LightingFixtureScheduleWarning>();
      TableAnalyzeExport template = LoadTemplateFromEmbeddedResource(out resourceName);
      ValidateTemplate(template, warnings);

      PromptPointResult pointResult = context.Editor.GetPoint(
        "\nSelect insertion point for lighting fixture schedule: "
      );
      if (pointResult.Status != PromptStatus.OK)
      {
        statusMessage = "Placement cancelled.";
        return null;
      }

      LightingFixtureScheduleSyncSchedule normalizedSchedule =
        NormalizeLightingFixtureScheduleForExternalUse(schedule);
      ObjectId newTableId = ObjectId.Null;
      string handle = "UNKNOWN";
      LightingFixtureScheduleVisualNormalizationResult normalizationResult = null;
      DocumentLock docLock = null;

      try
      {
        if (context.Document != null)
        {
          docLock = context.Document.LockDocument();
        }

        using (Transaction tr = context.Database.TransactionManager.StartTransaction())
        {
          Table table = RecreateTable(
            template,
            context.Database,
            tr,
            pointResult.Value,
            warnings
          );

          BlockTableRecord currentSpace =
            tr.GetObject(context.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          if (currentSpace == null)
          {
            throw new InvalidOperationException(
              "Unable to open current space for table insertion."
            );
          }

          newTableId = currentSpace.AppendEntity(table);
          tr.AddNewlyCreatedDBObject(table, true);

          normalizationResult = ApplyLightingScheduleToTable(
            table,
            context.Database,
            tr,
            normalizedSchedule
          );
          SafeApply(
            warnings,
            "table.generateLayout",
            () => table.GenerateLayout()
          );
          SafeApply(
            warnings,
            "table.recomputeTableBlock",
            () => TryRecomputeTableBlock(table)
          );

          handle = SafeGet(
            warnings,
            "table.handle",
            () => table.Handle.ToString(),
            null,
            null,
            "UNKNOWN"
          );

          tr.Commit();
        }
      }
      finally
      {
        docLock?.Dispose();
      }

      LightingFixtureScheduleStoreRecord record = SaveLightingFixtureScheduleStoreRecord(
        context.ProjectId,
        normalizedSchedule,
        context.DwgPath,
        handle,
        LightingFixtureScheduleStoreUpdatedByAutoCAD,
        loadedVersion > 0 ? (long?)loadedVersion : null,
        loadedVersion > 0 ? (long?)loadedVersion : null
      );

      context.TableId = newTableId;
      context.TableHandle = handle;
      SaveLightingFixtureScheduleStoreLink(
        context.ProjectId,
        context.DwgPath,
        handle,
        record?.Version ?? 0
      );
      SafeEditorRegen(context.Editor);

      string versionMessage = record == null
        ? "Store update unavailable."
        : record.HadConflict
          ? $"Saved version {record.Version}; a newer central version was overwritten."
          : $"Saved version {record.Version}.";
      string normalizationMessage = normalizationResult == null
        ? "Visual normalization unavailable."
        : $"Visual normalization: {FormatNormalizationSummary(normalizationResult)}.";
      statusMessage = $"Placed and linked table {handle}. {versionMessage} {normalizationMessage}";
      if (warnings.Count > 0)
      {
        statusMessage += $" Warnings: {warnings.Count}.";
      }

      return record;
    }

    private static void LightingFixtureScheduleAutoSyncOnIdle(object sender, EventArgs e)
    {
      if (s_lightingFixtureScheduleAutoSyncBusy || DateTime.UtcNow < s_nextLightingFixtureScheduleSyncUtc)
      {
        return;
      }

      s_nextLightingFixtureScheduleSyncUtc = DateTime.UtcNow.AddSeconds(2);
      Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      try
      {
        s_lightingFixtureScheduleAutoSyncBusy = true;
        TryApplyLightingFixtureScheduleStoreUpdateToActiveDocument(doc);
      }
      catch
      {
      }
      finally
      {
        s_lightingFixtureScheduleAutoSyncBusy = false;
      }
    }

    private static void TryApplyLightingFixtureScheduleStoreUpdateToActiveDocument(Document doc)
    {
      if (doc == null)
      {
        return;
      }

      string dwgPath = ResolveBestDwgPath(doc, doc.Database);
      if (string.IsNullOrWhiteSpace(dwgPath))
      {
        return;
      }

      LightingFixtureScheduleStoreLink link =
        GetLightingFixtureScheduleStoreLinkByDwgPath(dwgPath);
      if (link == null || string.IsNullOrWhiteSpace(link.ProjectId) || string.IsNullOrWhiteSpace(link.TableHandle))
      {
        return;
      }

      LightingFixtureScheduleStoreRecord record =
        GetLightingFixtureScheduleStoreRecord(link.ProjectId);
      if (record == null || record.Version <= link.LastAppliedVersion)
      {
        return;
      }

      var context = new LightingFixtureScheduleBindingContext
      {
        Document = doc,
        Database = doc.Database,
        Editor = doc.Editor,
        ProjectId = link.ProjectId,
        DwgPath = dwgPath,
        TableHandle = link.TableHandle,
        TableId = ResolveLightingFixtureScheduleTableIdByHandle(doc.Database, link.TableHandle),
      };

      ApplyLightingFixtureScheduleRecordToLinkedTable(context, record);
      if (
        IsLightingFixtureScheduleEditorControlReady() &&
        s_lightingFixtureSchedulePaletteContext != null &&
        string.Equals(s_lightingFixtureSchedulePaletteContext.ProjectId, record.ProjectId, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(s_lightingFixtureSchedulePaletteContext.DwgPath, dwgPath, StringComparison.OrdinalIgnoreCase)
      )
      {
        if (!s_lightingFixtureScheduleEditorControl.HasPendingLocalEdits)
        {
          s_lightingFixtureSchedulePaletteContext.TableHandle = context.TableHandle;
          s_lightingFixtureSchedulePaletteContext.TableId = context.TableId;
          s_lightingFixtureScheduleEditorControl.LoadRecord(
            s_lightingFixtureSchedulePaletteContext,
            record,
            $"Applied central version {record.Version} to the linked AutoCAD table."
          );
        }
        else
        {
          s_lightingFixtureScheduleEditorControl.MarkExternalUpdateAvailable(record.Version);
        }
      }
    }

    private static void ApplyLightingFixtureScheduleRecordToLinkedTable(
      LightingFixtureScheduleBindingContext context,
      LightingFixtureScheduleStoreRecord record
    )
    {
      if (context == null || record == null || context.Database == null)
      {
        return;
      }

      if (string.IsNullOrWhiteSpace(context.TableHandle) && context.TableId != ObjectId.Null)
      {
        using (Transaction tr = context.Database.TransactionManager.StartTransaction())
        {
          Table handleTable = tr.GetObject(context.TableId, OpenMode.ForRead, false) as Table;
          if (handleTable != null)
          {
            context.TableHandle = handleTable.Handle.ToString();
          }
          tr.Commit();
        }
      }

      if (context.TableId == ObjectId.Null)
      {
        context.TableId = ResolveLightingFixtureScheduleTableIdByHandle(
          context.Database,
          context.TableHandle
        );
      }
      if (context.TableId == ObjectId.Null)
      {
        throw new InvalidOperationException("No linked AutoCAD table is available. Link a table first.");
      }

      if (context.Document != null)
      {
        using (DocumentLock docLock = context.Document.LockDocument())
        using (Transaction tr = context.Database.TransactionManager.StartTransaction())
        {
          Table table = tr.GetObject(context.TableId, OpenMode.ForWrite, false) as Table;
          if (table == null)
          {
            throw new InvalidOperationException("Linked AutoCAD table could not be opened.");
          }

          LightingFixtureScheduleVisualNormalizationResult normalizationResult =
            ApplyLightingScheduleToTable(table, context.Database, tr, record.Schedule);
          table.GenerateLayout();
          TryRecomputeTableBlock(table);
          context.TableHandle = table.Handle.ToString();
          tr.Commit();

          SaveLightingFixtureScheduleStoreLink(
            context.ProjectId,
            context.DwgPath,
            context.TableHandle,
            record.Version
          );
          SafeEditorRegen(context.Editor);
          if (IsLightingFixtureScheduleEditorControlReady())
          {
            s_lightingFixtureScheduleEditorControl.SetStatus(
              $"Applied table update. Visual normalization: {FormatNormalizationSummary(normalizationResult)}"
            );
          }
        }
      }
    }
  }

  internal interface ILightingFixtureScheduleHost : IDisposable
  {
    string HostKind { get; }
    bool IsDisposed { get; }
    bool IsVisible { get; }
    void Show();
    void Hide();
  }

  internal sealed class LightingFixtureSchedulePaletteDockHost : ILightingFixtureScheduleHost
  {
    private PaletteSet _palette;

    internal LightingFixtureSchedulePaletteDockHost(Control control, Guid paletteGuid)
    {
      if (control == null)
      {
        throw new ArgumentNullException(nameof(control));
      }

      _palette = new PaletteSet("Lighting Fixture Schedule", paletteGuid);
      try
      {
        _palette.Style =
          PaletteSetStyles.ShowPropertiesMenu |
          PaletteSetStyles.ShowAutoHideButton |
          PaletteSetStyles.ShowCloseButton;
      }
      catch
      {
      }

      try
      {
        _palette.DockEnabled = DockSides.Left | DockSides.Right;
      }
      catch
      {
      }

      try
      {
        _palette.MinimumSize = new Size(720, 480);
      }
      catch
      {
      }

      try
      {
        _palette.Add("Schedule", control);
      }
      catch
      {
        Dispose();
        throw;
      }
    }

    public string HostKind => "AutoCAD palette";

    public bool IsDisposed => _palette == null;

    public bool IsVisible => _palette != null && _palette.Visible;

    public void Show()
    {
      EnsurePalette();
      _palette.Visible = true;
    }

    public void Hide()
    {
      if (_palette != null)
      {
        _palette.Visible = false;
      }
    }

    public void Dispose()
    {
      if (_palette == null)
      {
        return;
      }

      try
      {
        _palette.Visible = false;
      }
      catch
      {
      }

      _palette.Dispose();
      _palette = null;
    }

    private void EnsurePalette()
    {
      if (_palette == null)
      {
        throw new ObjectDisposedException(nameof(LightingFixtureSchedulePaletteDockHost));
      }
    }
  }

  internal sealed class LightingFixtureScheduleModelessFormHost : ILightingFixtureScheduleHost
  {
    private readonly LightingFixtureScheduleModelessForm _form;
    private bool _shownOnce;

    internal LightingFixtureScheduleModelessFormHost(Control control)
    {
      if (control == null)
      {
        throw new ArgumentNullException(nameof(control));
      }

      _form = new LightingFixtureScheduleModelessForm(control);
    }

    public string HostKind => "modeless WinForms window";

    public bool IsDisposed => _form == null || _form.IsDisposed;

    public bool IsVisible => _form != null && !_form.IsDisposed && _form.Visible;

    public void Show()
    {
      EnsureForm();

      if (!_shownOnce)
      {
        try
        {
          Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(_form);
        }
        catch
        {
          _form.Show();
        }

        _shownOnce = true;
      }
      else if (!_form.Visible)
      {
        _form.Show();
      }

      _form.BringToFront();
      _form.Focus();
    }

    public void Hide()
    {
      if (_form != null && !_form.IsDisposed)
      {
        _form.Hide();
      }
    }

    public void Dispose()
    {
      if (_form == null || _form.IsDisposed)
      {
        return;
      }

      _form.AllowDispose = true;
      _form.Close();
      _form.Dispose();
    }

    private void EnsureForm()
    {
      if (_form == null || _form.IsDisposed)
      {
        throw new ObjectDisposedException(nameof(LightingFixtureScheduleModelessFormHost));
      }
    }
  }

  internal sealed class LightingFixtureScheduleModelessForm : Form
  {
    internal LightingFixtureScheduleModelessForm(Control content)
    {
      if (content == null)
      {
        throw new ArgumentNullException(nameof(content));
      }

      Text = "Lighting Fixture Schedule";
      ShowInTaskbar = false;
      StartPosition = FormStartPosition.CenterScreen;
      FormBorderStyle = FormBorderStyle.SizableToolWindow;
      MinimumSize = new Size(720, 480);
      Size = new Size(1024, 720);

      content.Dock = DockStyle.Fill;
      Controls.Add(content);
    }

    internal bool AllowDispose { get; set; }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
      if (!AllowDispose && e.CloseReason == CloseReason.UserClosing)
      {
        e.Cancel = true;
        Hide();
        return;
      }

      base.OnFormClosing(e);
    }
  }
}
