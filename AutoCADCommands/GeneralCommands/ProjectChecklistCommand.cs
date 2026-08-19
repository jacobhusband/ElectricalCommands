using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Windows;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private static ProjectChecklistWindow _activeChecklistWindow;

    [CommandMethod("CHECKLIST", CommandFlags.Session)]
    [CommandMethod("CHECKLISTS", CommandFlags.Session)]
    [CommandMethod("PROJECTCHECKLIST", CommandFlags.Session)]
    [CommandMethod("PCL", CommandFlags.Session)]
    public void ProjectChecklist()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        MessageBox.Show("No active AutoCAD document is available.", "Checklists", MessageBoxButton.OK, MessageBoxImage.Warning);
        return;
      }

      var db = doc.Database;
      var ed = doc.Editor;

      try
      {
        string dwgPath = ProjectChecklistStore.ResolveBestDwgPath(doc, db);
        if (string.IsNullOrWhiteSpace(dwgPath))
        {
          ed.WriteMessage("\nCHECKLIST: Please save the active drawing to a project folder first.");
          MessageBox.Show(
            "Please save the active drawing first so checklists can be stored and shared within the drawing's project folder.",
            "Save Drawing Required",
            MessageBoxButton.OK,
            MessageBoxImage.Information
          );
          return;
        }

        string folderPath = ProjectChecklistStore.ResolveDrawingFolder(dwgPath);
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
          ed.WriteMessage("\nCHECKLIST: Unable to determine the project directory for the active drawing.");
          return;
        }

        if (_activeChecklistWindow != null && _activeChecklistWindow.IsLoaded)
        {
          _activeChecklistWindow.SwitchFolderOrDrawing(folderPath, dwgPath);
          if (_activeChecklistWindow.WindowState == WindowState.Minimized)
          {
            _activeChecklistWindow.WindowState = WindowState.Normal;
          }
          _activeChecklistWindow.Activate();
          _activeChecklistWindow.Focus();
          ed.WriteMessage($"\nCHECKLIST: Activated project checklist window for {folderPath}.");
          return;
        }

        var definitions = ProjectChecklistStore.LoadAllChecklists(folderPath);
        var state = ProjectChecklistStore.LoadFolderState(folderPath);

        ed.WriteMessage($"\nCHECKLIST: Opening modeless project checklists for {folderPath}...");

        _activeChecklistWindow = new ProjectChecklistWindow(folderPath, dwgPath, definitions, state);
        _activeChecklistWindow.Closed += (s, e) => { _activeChecklistWindow = null; };

        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(_activeChecklistWindow);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nCHECKLIST error: {ex.Message}");
      }
    }
  }
}
