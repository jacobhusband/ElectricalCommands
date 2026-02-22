using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("TXTINCR", CommandFlags.UsePickSet)]
    [CommandMethod("TI", CommandFlags.UsePickSet)]
    public void Incrementer()
    {
      Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      try
      {
        // Get user inputs
        string prefix = ed.GetString("\nEnter prefix (e.g., HP-): ").StringResult;
        int startNum = Convert.ToInt32(ed.GetString("\nEnter start number: ").StringResult);
        int endNum = Convert.ToInt32(ed.GetString("\nEnter end number: ").StringResult);
        string oddEven = ed.GetString("\nEnter 'O' for odd or 'E' for even: ").StringResult.ToUpper();

        // Handle selection
        SelectionSet sset;
        PromptSelectionResult selRes = ed.SelectImplied();
        if (selRes.Status == PromptStatus.OK)
        {
          // Use the PickFirst selection
          sset = selRes.Value;
        }
        else
        {
          // If no PickFirst selection, prompt for selection
          PromptSelectionOptions pso = new PromptSelectionOptions();
          pso.MessageForAdding = "\nSelect MText and DBText objects: ";
          TypedValue[] filterList = new TypedValue[]
          {
                new TypedValue((int)DxfCode.Start, "TEXT,MTEXT")
          };
          SelectionFilter filter = new SelectionFilter(filterList);
          selRes = ed.GetSelection(pso, filter);
          if (selRes.Status != PromptStatus.OK)
            return;
          sset = selRes.Value;
        }

        // Get point from user
        PromptPointResult ppr = ed.GetPoint("\nSelect a reference point: ");
        if (ppr.Status != PromptStatus.OK) return;
        Point3d selectedPoint = ppr.Value;

        // Collect selected MText and DBText objects
        List<(Entity entity, double distance)> textObjects = new List<(Entity, double)>();
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId objId in sset.GetObjectIds())
          {
            Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
            if (ent is MText || ent is DBText)
            {
              double dist = ent.GeometricExtents.MinPoint.DistanceTo(selectedPoint);
              textObjects.Add((ent, dist));
            }
          }

          // Sort objects by distance
          textObjects = textObjects.OrderBy(x => x.distance).ToList();

          // Update text objects
          int currentNum = startNum;
          foreach (var (ent, _) in textObjects)
          {
            ent.UpgradeOpen();
            string newText = $"{prefix}{currentNum}";
            if (ent is MText mtext)
            {
              mtext.Contents = newText;
            }
            else if (ent is DBText dbtext)
            {
              dbtext.TextString = newText;
            }

            if (currentNum < endNum)
            {
              do
              {
                currentNum++;
              } while ((oddEven == "O" && currentNum % 2 == 0) || (oddEven == "E" && currentNum % 2 != 0));
              if (currentNum > endNum)
              {
                currentNum = endNum;
              }
            }
          }

          tr.Commit();
        }

        ed.WriteMessage($"\nINCREMENTER command completed successfully. Modified {textObjects.Count} text object(s).");

        // Clear the PickFirst selection set
        ed.SetImpliedSelection(new ObjectId[0]);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nError: {ex.Message}");
      }
    }
  }
}
