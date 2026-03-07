using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
  public static class TextSelectCommand
  {
    [CommandMethod("TEXTSELECT", CommandFlags.UsePickSet)]
    public static void TextSelect()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      try
      {
        PromptSelectionResult selection = ed.SelectImplied();
        if (selection.Status != PromptStatus.OK || selection.Value == null || selection.Value.Count == 0)
        {
          PromptSelectionOptions options = new PromptSelectionOptions();
          options.MessageForAdding = "\nSelect objects: ";
          selection = ed.GetSelection(options);
        }

        if (selection.Status != PromptStatus.OK || selection.Value == null || selection.Value.Count == 0)
        {
          ed.WriteMessage("\nNothing selected.");
          return;
        }

        List<ObjectId> textIds = new List<ObjectId>();
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId objectId in selection.Value.GetObjectIds())
          {
            DBObject obj = tr.GetObject(objectId, OpenMode.ForRead, false);
            if (obj is DBText || obj is MText)
            {
              textIds.Add(objectId);
            }
          }

          tr.Commit();
        }

        if (textIds.Count == 0)
        {
          ed.SetImpliedSelection(Array.Empty<ObjectId>());
          ed.WriteMessage("\nNo TEXT or MTEXT objects found in selection.");
          return;
        }

        ed.SetImpliedSelection(textIds.ToArray());
        ed.WriteMessage($"\nSelected {textIds.Count} text object(s).");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nTEXTSELECT error: {ex.Message}");
      }
    }
  }
}
