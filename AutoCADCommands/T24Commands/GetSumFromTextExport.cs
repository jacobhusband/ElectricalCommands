using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("TXTSUMEXPORT", CommandFlags.UsePickSet)]
    [CommandMethod("TSE", CommandFlags.UsePickSet)]
    public void SumTextExport()
    {
      Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      try
      {
        // 1. Check for implied selection (PickFirst)
        PromptSelectionResult psr = ed.SelectImplied();
        if (psr.Status != PromptStatus.OK)
        {
          // If no implied selection, prompt the user to select objects
          PromptSelectionOptions pso = new PromptSelectionOptions();
          pso.MessageForAdding = "Select DBText or MText objects: ";
          TypedValue[] filterList = new TypedValue[]
          {
              new TypedValue((int)DxfCode.Start, "TEXT,MTEXT")
          };
          SelectionFilter filter = new SelectionFilter(filterList);
          psr = ed.GetSelection(pso, filter);
        }

        if (psr.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo objects selected.");
          return;
        }

        List<RoomInfo> roomInfoList = new List<RoomInfo>();
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          // Process selected objects
          SelectionSet ss = psr.Value;
          foreach (SelectedObject so in ss)
          {
            Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
            if (ent is DBText || ent is MText)
            {
              string text = (ent is DBText) ? ((DBText)ent).TextString : ((MText)ent).Text;
              Point3d position = (ent is DBText) ? ((DBText)ent).Position : ((MText)ent).Location;
              if (text.Contains("sq ft"))
              {
                // Find nearest text object
                Entity nearestEnt = FindNearestTextObject(tr, ss, ent, position);
                if (nearestEnt != null)
                {
                  string roomType = (nearestEnt is DBText) ? ((DBText)nearestEnt).TextString : ((MText)nearestEnt).Text;
                  double squareFeet = ExtractSquareFeet(text);
                  roomInfoList.Add(new RoomInfo { RoomType = roomType, SquareFeet = squareFeet });
                }
              }
            }
          }
          tr.Commit();
        }
        // 5. Combine similar room types
        var combinedRooms = roomInfoList
            .GroupBy(r => r.RoomType)
            .Select(g => new RoomInfo
            {
              RoomType = g.Key,
              SquareFeet = g.Sum(r => r.SquareFeet)
            })
            .ToList();
        // 6. Output to JSON file
        string dwgPath = db.Filename;
        string jsonPath = Path.Combine(Path.GetDirectoryName(dwgPath), "T24Output.json");
        string json = JsonConvert.SerializeObject(combinedRooms, Formatting.Indented);
        File.WriteAllText(jsonPath, json);
        ed.WriteMessage($"\nExported room information to: {jsonPath}");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nError: {ex.Message}");
      }
      finally
      {
        // Clear the PickFirst selection set
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    private Entity FindNearestTextObject(Transaction tr, SelectionSet ss, Entity currentEnt, Point3d position)
    {
      Entity nearestEnt = null;
      double minDistance = double.MaxValue;

      foreach (SelectedObject so in ss)
      {
        Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
        if ((ent is DBText || ent is MText) && ent != currentEnt)
        {
          Point3d entPosition = (ent is DBText) ? ((DBText)ent).Position : ((MText)ent).Location;
          double distance = position.DistanceTo(entPosition);
          if (distance < minDistance)
          {
            minDistance = distance;
            nearestEnt = ent;
          }
        }
      }

      return nearestEnt;
    }

    private double ExtractSquareFeet(string text)
    {
      string[] parts = text.Split(' ');
      if (parts.Length > 0 && double.TryParse(parts[0], out double result))
      {
        return result;
      }
      return 0;
    }
  }

  public class RoomInfo
  {
    public string RoomType { get; set; }
    public double SquareFeet { get; set; }
  }
}
