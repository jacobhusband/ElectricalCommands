using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("TXTNEW", CommandFlags.UsePickSet)]
    [CommandMethod("TN", CommandFlags.UsePickSet)]
    public void TextNew()
    {
      Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      try
      {
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
          pso.MessageForAdding = "Select text objects to modify:";
          TypedValue[] filterList = new TypedValue[]
          {
                new TypedValue((int)DxfCode.Start, "TEXT,MTEXT")
          };
          SelectionFilter filter = new SelectionFilter(filterList);
          selRes = ed.GetSelection(pso, filter);
          if (selRes.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nCommand canceled.");
            return;
          }
          sset = selRes.Value;
        }

        // Filter for TEXT and MTEXT objects
        ObjectId[] filteredIds;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          filteredIds = sset.GetObjectIds().Where(id =>
          {
            var obj = tr.GetObject(id, OpenMode.ForRead);
            return obj is DBText || obj is MText;
          }).ToArray();
          tr.Commit();
        }

        if (filteredIds.Length == 0)
        {
          ed.WriteMessage("\nNo text objects selected.");
          return;
        }

        // Prompt user for new text content
        PromptResult pr = ed.GetString("\nEnter new text content: ");
        if (pr.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nCommand canceled.");
          return;
        }
        string newContent = pr.StringResult;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          int processedCount = 0;
          foreach (ObjectId objId in filteredIds)
          {
            Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
            if (ent is DBText text)
            {
              text.TextString = newContent;
              processedCount++;
            }
            else if (ent is MText mtext)
            {
              mtext.Contents = newContent;
              processedCount++;
            }
          }
          tr.Commit();
          ed.WriteMessage($"\nCommand completed successfully. Modified {processedCount} text object(s).");
        }

        // Clear the PickFirst selection set
        ed.SetImpliedSelection(new ObjectId[0]);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("\nError: " + ex.Message);
      }
    }
  }
}
