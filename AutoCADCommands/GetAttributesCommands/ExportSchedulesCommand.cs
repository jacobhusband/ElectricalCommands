using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("EXPORTSCHEDULES")]
    public void ExportSchedules()
    {
      ExecuteExportSchedules(null);
    }

    [CommandMethod("-EXPORTSCHEDULES")]
    public void ExportSchedulesNonInteractive()
    {
      var (_, _, ed) = Globals.GetGlobals();
      if (ed == null) return;

      var promptString = ed.GetString(new PromptStringOptions("\nEnter export JSON output path (or press Enter for default): "));
      string targetPath = promptString.Status == PromptStatus.OK && !string.IsNullOrWhiteSpace(promptString.StringResult)
          ? promptString.StringResult.Trim()
          : null;

      ExecuteExportSchedules(targetPath);
    }

    private void ExecuteExportSchedules(string explicitOutputPath)
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        ed?.WriteMessage("\nEXPORTSCHEDULES: No active AutoCAD document.");
        return;
      }

      try
      {
        var oleRecords = new List<OleFrameExportInfo>();
        var tableRecords = new List<TableExportInfo>();

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

          foreach (ObjectId btrId in bt)
          {
            var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            string spaceName = btr.IsLayout ? btr.Name : (btr.IsAnonymous ? "*Anonymous" : btr.Name);

            foreach (ObjectId entId in btr)
            {
              var ent = tr.GetObject(entId, OpenMode.ForRead, false);

              // 1. Ole2Frame entities (Excel schedules embedded or linked)
              if (ent is Ole2Frame ole)
              {
                oleRecords.Add(new OleFrameExportInfo
                {
                  Handle = ole.Handle.ToString(),
                  Layout = spaceName,
                  UserType = ole.UserType,
                  Type = ole.Type.ToString(),
                  MinPoint = ole.GeometricExtents.MinPoint != null
                      ? new double[] { ole.GeometricExtents.MinPoint.X, ole.GeometricExtents.MinPoint.Y, ole.GeometricExtents.MinPoint.Z }
                      : null,
                  MaxPoint = ole.GeometricExtents.MaxPoint != null
                      ? new double[] { ole.GeometricExtents.MaxPoint.X, ole.GeometricExtents.MaxPoint.Y, ole.GeometricExtents.MaxPoint.Z }
                      : null
                });
              }
              // 2. Native AutoCAD Table entities
              else if (ent is Table tbl)
              {
                string titleText = "";
                if (tbl.Rows.Count > 0 && tbl.Columns.Count > 0)
                {
                  try
                  {
                    titleText = tbl.Cells[0, 0].TextString ?? "";
                  }
                  catch
                  {
                    titleText = "";
                  }
                }

                tableRecords.Add(new TableExportInfo
                {
                  Handle = tbl.Handle.ToString(),
                  Layout = spaceName,
                  RowCount = tbl.Rows.Count,
                  ColumnCount = tbl.Columns.Count,
                  TableStyleName = tbl.TableStyle.IsValid ? ((TableStyle)tr.GetObject(tbl.TableStyle, OpenMode.ForRead)).Name : "Standard",
                  Title = titleText
                });
              }
            }
          }

          tr.Commit();
        }

        // Determine destination output file path
        string outputPath = explicitOutputPath;
        if (string.IsNullOrWhiteSpace(outputPath))
        {
          string dwgName = Path.GetFileNameWithoutExtension(doc.Name);
          string dwgDir = Path.GetDirectoryName(doc.Name);
          if (string.IsNullOrEmpty(dwgDir))
          {
            dwgDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
          }
          string exportsDir = Path.Combine(dwgDir, "exports");
          Directory.CreateDirectory(exportsDir);
          outputPath = Path.Combine(exportsDir, $"{dwgName}_Schedules.json");
        }
        else
        {
          string targetDir = Path.GetDirectoryName(outputPath);
          if (!string.IsNullOrEmpty(targetDir))
          {
            Directory.CreateDirectory(targetDir);
          }
        }

        var exportPayload = new ScheduleExportPayload
        {
          DrawingPath = doc.Name,
          ExportTimestamp = DateTime.UtcNow.ToString("o"),
          OleFrames = oleRecords,
          Tables = tableRecords,
          TotalOleCount = oleRecords.Count,
          TotalTableCount = tableRecords.Count,
          Status = "SUCCESS"
        };

        string jsonContent = JsonConvert.SerializeObject(exportPayload, Formatting.Indented);
        File.WriteAllText(outputPath, jsonContent);

        ed.WriteMessage($"\nEXPORTSCHEDULES: Exported {oleRecords.Count} OLE schedules and {tableRecords.Count} native tables to: {outputPath}");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nEXPORTSCHEDULES Error: {ex.Message}");
      }
    }
  }

  public class ScheduleExportPayload
  {
    public string DrawingPath { get; set; }
    public string ExportTimestamp { get; set; }
    public List<OleFrameExportInfo> OleFrames { get; set; } = new List<OleFrameExportInfo>();
    public List<TableExportInfo> Tables { get; set; } = new List<TableExportInfo>();
    public int TotalOleCount { get; set; }
    public int TotalTableCount { get; set; }
    public string Status { get; set; }
  }

  public class OleFrameExportInfo
  {
    public string Handle { get; set; }
    public string Layout { get; set; }
    public string UserType { get; set; }
    public string Type { get; set; }
    public double[] MinPoint { get; set; }
    public double[] MaxPoint { get; set; }
  }

  public class TableExportInfo
  {
    public string Handle { get; set; }
    public string Layout { get; set; }
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public string TableStyleName { get; set; }
    public string Title { get; set; }
  }
}
