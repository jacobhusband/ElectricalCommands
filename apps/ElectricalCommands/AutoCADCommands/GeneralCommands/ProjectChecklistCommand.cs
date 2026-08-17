using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("PROJECTCHECKLIST", CommandFlags.Modal)]
    [CommandMethod("PCL", CommandFlags.Modal)]
    public void ProjectChecklist()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      string stage = "initialization";
      try
      {
        stage = "drawing path resolution";
        string dwgPath = ProjectChecklistStore.RequireSavedDwgPath(doc, db);

        stage = "project binding resolution";
        ProjectChecklistBindingContext context =
          ResolveProjectChecklistBindingContext(doc, db, ed, dwgPath);

        stage = "shared store load";
        ProjectChecklistStoreRecord record =
          ProjectChecklistStore.EnsureRecord(context.ProjectId, context.DwgPath);

        stage = "checklist definition load";
        ProjectChecklistDefinitionsLoadResult definitionsResult =
          ProjectChecklistStore.LoadChecklistDefinitionsResult();
        List<ProjectChecklistDefinition> definitions = definitionsResult.Definitions;

        stage = "dialog creation";
        using (var editorControl = new ProjectChecklistPaletteControl())
        {
          EventHandler<ProjectChecklistSaveRequestedEventArgs> saveHandler =
            (sender, e) => SaveProjectChecklist(context, editorControl, e);
          EventHandler reloadHandler =
            (sender, e) => ReloadProjectChecklist(context, editorControl);
          editorControl.SaveRequested += saveHandler;
          editorControl.ReloadRequested += reloadHandler;

          try
          {
            stage = "record load";
            editorControl.LoadRecord(
              context,
              definitions,
              record,
              definitions.Count == 0
                ? definitionsResult.StatusMessage
                : "Editor open. Checkbox changes save to the shared project checklist store."
            );

            stage = "dialog show";
            var window = new ProjectChecklistWindow(editorControl);
            window.ShowDialog();
            ed.WriteMessage("\nPCL: project checklist dialog closed.");
          }
          finally
          {
            editorControl.SaveRequested -= saveHandler;
            editorControl.ReloadRequested -= reloadHandler;
          }
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nPCL error during {stage}: {ex.Message}");
      }
    }

    private static ProjectChecklistBindingContext ResolveProjectChecklistBindingContext(
      Document doc,
      Database db,
      Editor ed,
      string dwgPath
    )
    {
      string projectId = ProjectChecklistStore.ResolveProjectIdForDrawing(dwgPath);
      if (string.IsNullOrWhiteSpace(projectId))
      {
        projectId = PromptForProjectChecklistProjectId(ed);
      }
      if (string.IsNullOrWhiteSpace(projectId))
      {
        throw new InvalidOperationException(
          "A project ID is required before project checklists can use the shared store."
        );
      }

      ProjectChecklistStore.SaveLink(projectId, dwgPath);
      return new ProjectChecklistBindingContext
      {
        Document = doc,
        Database = db,
        Editor = ed,
        ProjectId = projectId,
        DwgPath = dwgPath,
      };
    }

    private static string PromptForProjectChecklistProjectId(Editor ed)
    {
      if (ed == null)
      {
        return string.Empty;
      }

      PromptStringOptions options = new PromptStringOptions(
        "\nEnter project ID to bind project checklists: "
      );
      options.AllowSpaces = true;
      PromptResult result = ed.GetString(options);
      return result.Status == PromptStatus.OK
        ? ProjectChecklistStore.NormalizePlainText(result.StringResult)
        : string.Empty;
    }

    private static void SaveProjectChecklist(
      ProjectChecklistBindingContext context,
      IProjectChecklistEditorControl editorControl,
      ProjectChecklistSaveRequestedEventArgs e
    )
    {
      if (context == null || editorControl == null)
      {
        return;
      }

      try
      {
        ProjectChecklistStoreRecord savedRecord = ProjectChecklistStore.SaveRecord(
          context.ProjectId,
          e.State,
          context.DwgPath,
          "autocad",
          e.LoadedVersion
        );
        ProjectChecklistDefinitionsLoadResult definitionsResult =
          ProjectChecklistStore.LoadChecklistDefinitionsResult();
        List<ProjectChecklistDefinition> definitions = definitionsResult.Definitions;
        if (IsProjectChecklistEditorControlReady(editorControl))
        {
          editorControl.LoadRecord(
            context,
            definitions,
            savedRecord,
            definitions.Count == 0
              ? definitionsResult.StatusMessage
              : savedRecord.HadConflict
              ? $"Saved version {savedRecord.Version}. A newer central version was overwritten."
              : $"Saved version {savedRecord.Version}."
          );
        }
      }
      catch (System.Exception ex)
      {
        SetProjectChecklistStatus(editorControl, $"Save failed: {ex.Message}");
      }
    }

    private static void ReloadProjectChecklist(
      ProjectChecklistBindingContext context,
      IProjectChecklistEditorControl editorControl
    )
    {
      if (context == null || editorControl == null)
      {
        return;
      }

      try
      {
        ProjectChecklistStoreRecord record =
          ProjectChecklistStore.EnsureRecord(context.ProjectId, context.DwgPath);
        ProjectChecklistDefinitionsLoadResult definitionsResult =
          ProjectChecklistStore.LoadChecklistDefinitionsResult();
        List<ProjectChecklistDefinition> definitions = definitionsResult.Definitions;
        if (IsProjectChecklistEditorControlReady(editorControl))
        {
          editorControl.LoadRecord(
            context,
            definitions,
            record,
            definitions.Count == 0
              ? definitionsResult.StatusMessage
              : $"Reloaded project checklist version {record.Version}."
          );
        }
      }
      catch (System.Exception ex)
      {
        SetProjectChecklistStatus(editorControl, $"Reload failed: {ex.Message}");
      }
    }

    private static bool IsProjectChecklistEditorControlReady(
      IProjectChecklistEditorControl editorControl
    )
    {
      return
        editorControl != null &&
        editorControl.HostControl != null &&
        !editorControl.HostControl.IsDisposed;
    }

    private static void SetProjectChecklistStatus(
      IProjectChecklistEditorControl editorControl,
      string message
    )
    {
      if (IsProjectChecklistEditorControlReady(editorControl))
      {
        editorControl.SetStatus(message);
      }
    }
  }
}
