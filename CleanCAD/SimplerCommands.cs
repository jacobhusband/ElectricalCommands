using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;

namespace AutoCADCleanupTool
{
    public class SimplerCommands
    {
        // The CommandFlags.UsePickSet flag tells AutoCAD that this command
        // is aware of and can use the pre-selected "PickFirst" set.
        [CommandMethod("EraseOther", CommandFlags.UsePickSet)]
        public static void EraseOtherCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptSelectionResult psr;

            // Step 1: Use SelectImplied() to get the pre-selected ("PickFirst") set.
            psr = ed.SelectImplied();

            // If the status is not OK, it means nothing was pre-selected.
            // In this case, we fall back to prompting the user for a new selection.
            if (psr.Status != PromptStatus.OK)
            {
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\nSelect objects to keep: ";
                psr = ed.GetSelection(pso);
            }

            // If the user cancelled or made an empty selection at any point, exit.
            if (psr.Status != PromptStatus.OK)
            {
                return;
            }

            // Create a HashSet for fast lookups of the objects we want to keep.
            var idsToKeep = new HashSet<ObjectId>(psr.Value.GetObjectIds());
            int erasedCount = 0;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    var layersToUnlock = new HashSet<ObjectId>();
                    var idsToErase = new List<ObjectId>();

                    // Iterate through the current space to find objects to erase.
                    foreach (ObjectId id in currentSpace)
                    {
                        if (idsToKeep.Contains(id))
                        {
                            continue; // Skip objects in our "keep" list.
                        }

                        Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent != null)
                        {
                            idsToErase.Add(id);
                            LayerTableRecord layer = trans.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                            if (layer != null && layer.IsLocked)
                            {
                                layersToUnlock.Add(ent.LayerId);
                            }
                        }
                    }

                    // Temporarily unlock any locked layers.
                    foreach (ObjectId layerId in layersToUnlock)
                    {
                        LayerTableRecord layer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        layer.IsLocked = false;
                    }

                    // Erase the objects.
                    foreach (ObjectId idToErase in idsToErase)
                    {
                        Entity entToErase = trans.GetObject(idToErase, OpenMode.ForWrite) as Entity;
                        entToErase.Erase();
                        erasedCount++;
                    }

                    // Re-lock the layers that we unlocked.
                    foreach (ObjectId layerId in layersToUnlock)
                    {
                        LayerTableRecord layer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        layer.IsLocked = true;
                    }

                    trans.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred: {ex.Message}");
                    trans.Abort();
                }
            }

            if (erasedCount > 0)
            {
                ed.WriteMessage($"\nErased {erasedCount} object(s).");
                ed.Regen();
            }
        }
    }
}