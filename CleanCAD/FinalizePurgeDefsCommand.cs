using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        // STAGE 3: Surgically remove the image definitions from our kill list.
        [CommandMethod("-FINALIZE-PURGEDEFS", CommandFlags.Modal)]
        public static void FinalizePurgeDefsCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;
        
            ed.WriteMessage("\n--- Stage 3: Surgically removing orphaned image definitions... ---");
            if (_imageDefsToPurge.Count == 0)
            {
                ed.WriteMessage("\nNo orphaned image definitions were identified. Finalization complete.");
                TriggerKeepOnlyTitleBlockIfRequested(doc, ed);
                return;
            }
        
            try
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    int defsPurged = 0;
                    DBDictionary namedObjectsDict = trans.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (namedObjectsDict.Contains("ACAD_IMAGE_DICT"))
                    {
                        DBDictionary imageDict = trans.GetObject(namedObjectsDict.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite) as DBDictionary;
                        var entriesToRemove = new Dictionary<string, ObjectId>();
        
                        // Find the dictionary keys for the defs on our kill list
                        foreach (DBDictionaryEntry entry in imageDict)
                        {
                            if (_imageDefsToPurge.Contains(entry.Value))
                            {
                                entriesToRemove.Add(entry.Key, entry.Value);
                            }
                        }
                        
                        ed.WriteMessage($"\nFound {entriesToRemove.Count} dictionary entr(ies) to remove.");
                        foreach (var item in entriesToRemove)
                        {
                            ed.WriteMessage($"\n  - Detaching and Erasing: {item.Key}");
                            imageDict.Remove(item.Key);
                            DBObject imageDef = trans.GetObject(item.Value, OpenMode.ForWrite);
                            imageDef.Erase();
                            defsPurged++;
                        }
                    }
                    ed.WriteMessage($"\nSuccessfully purged {defsPurged} image definition(s).");
                    trans.Commit();
                }
                ed.WriteMessage("\n-------------------------------------------------");
                ed.WriteMessage("\nFinalization Complete. Please save the drawing.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during final purge: {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                _imageDefsToPurge.Clear();
                TriggerKeepOnlyTitleBlockIfRequested(doc, ed);
            }
        }

        private static void TriggerKeepOnlyTitleBlockIfRequested(Document doc, Editor ed)
        {
            if (!RunKeepOnlyAfterFinalize) return;
            RunKeepOnlyAfterFinalize = false;
            try
            {
                doc.SendStringToExecute("_.KEEPONLYTITLEBLOCKMS ", true, false, false);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to queue KEEPONLYTITLEBLOCKMS: {ex.Message}");
            }
        }
    }
}
