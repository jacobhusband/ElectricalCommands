using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace ElectricalCommands
{
  public partial class AuditCommands
  {
    private static ProjectAuditWindow _activeAuditWindow;

    [CommandMethod("PREFLIGHT", CommandFlags.Session)]
    [CommandMethod("PFL", CommandFlags.Session)]
    [CommandMethod("QAQC", CommandFlags.Session)]
    [CommandMethod("PROJECTAUDIT", CommandFlags.Session)]
    [CommandMethod("SCOPECHECK", CommandFlags.Session)]
    [CommandMethod("DYNAMICCHECKLIST", CommandFlags.Session)]
    public void ProjectAudit()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        MessageBox.Show("No active AutoCAD document is available.", "Pre-Flight QA/QC Engine", MessageBoxButton.OK, MessageBoxImage.Warning);
        return;
      }

      var db = doc.Database;
      var ed = doc.Editor;

      try
      {
        string dwgPath = ResolveBestDwgPath(doc, db);
        if (string.IsNullOrWhiteSpace(dwgPath))
        {
          ed.WriteMessage("\nPREFLIGHT: Please save the active drawing to a project folder first.");
          MessageBox.Show(
            "Please save the active drawing first so the project scope questionnaire and QA/QC audit state can be stored in the project folder.",
            "Save Drawing Required",
            MessageBoxButton.OK,
            MessageBoxImage.Information
          );
          return;
        }

        string folderPath = ResolveDrawingFolder(dwgPath);
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
          ed.WriteMessage("\nPREFLIGHT: Unable to determine the project directory for the active drawing.");
          return;
        }

        // If already open, activate and update drawing
        if (_activeAuditWindow != null && _activeAuditWindow.IsLoaded)
        {
          _activeAuditWindow.SwitchFolderOrDrawing(folderPath, dwgPath);
          if (_activeAuditWindow.WindowState == WindowState.Minimized)
          {
            _activeAuditWindow.WindowState = WindowState.Normal;
          }
          _activeAuditWindow.Activate();
          _activeAuditWindow.Focus();
          ed.WriteMessage($"\nPREFLIGHT: Activated QA/QC engine for {folderPath}.");
          return;
        }

        var masterCatalog = AuditEngine.GetMasterCatalog();
        var state = AuditEngine.LoadAuditState(folderPath);

        ed.WriteMessage($"\nPREFLIGHT: Opening dynamic pre-flight QA/QC engine for {folderPath}...");

        _activeAuditWindow = new ProjectAuditWindow(folderPath, dwgPath, masterCatalog, state);
        _activeAuditWindow.Closed += (s, e) => { _activeAuditWindow = null; };

        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(_activeAuditWindow);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nPREFLIGHT error: {ex.Message}");
      }
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

      return candidates
        .Select(p => p?.Trim())
        .Where(p => !string.IsNullOrWhiteSpace(p))
        .Where(p => string.Equals(Path.GetExtension(p), ".dwg", StringComparison.OrdinalIgnoreCase))
        .FirstOrDefault(p => !p.IndexOf("\\AppData\\Local\\Temp\\", StringComparison.OrdinalIgnoreCase).Equals(-1))
        ?? candidates
          .Select(p => p?.Trim())
          .Where(p => !string.IsNullOrWhiteSpace(p))
          .FirstOrDefault(p => string.Equals(Path.GetExtension(p), ".dwg", StringComparison.OrdinalIgnoreCase))
        ?? string.Empty;
    }

    private static string ResolveDrawingFolder(string dwgPath)
    {
      if (string.IsNullOrWhiteSpace(dwgPath)) return string.Empty;
      try
      {
        return Path.GetDirectoryName(dwgPath) ?? string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }
  }
}
