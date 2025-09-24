using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        // STAGE 1: Get snapshots, bind, and queue the main cleanup command.
        [CommandMethod("FINALIZE")]
        public static void FinalizeDrawingCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            SimplerCommands.EnsureAllLayersVisibleAndUnlocked(db, ed);
            SimplerCommands.PrepareXrefLayersForCleanup(db, ed);

            _blockIdsBeforeBind.Clear();
            _originalXrefIds.Clear();
            _imageDefsToPurge.Clear();

            SimplerCommands.DetachSpecialXrefs();

            ed.WriteMessage("\n--- Stage 1: Analyzing and Binding... ---");

            try
            {
                int bindCount = 0;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    foreach (ObjectId btrId in bt)
                    {
                        _blockIdsBeforeBind.Add(btrId);
                        var btr = trans.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr == null || !btr.IsFromExternalReference || !btr.IsResolved)
                        {
                            continue;
                        }

                        bool isDwg = false;
                        string pathName = btr.PathName ?? string.Empty;
                        if (!string.IsNullOrEmpty(pathName))
                        {
                            try { isDwg = string.Equals(Path.GetExtension(pathName), ".dwg", StringComparison.OrdinalIgnoreCase); }
                            catch { }
                        }

                        if (!isDwg)
                        {
                            string name = btr.Name ?? string.Empty;
                            isDwg = name.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase);
                        }

                        if (isDwg)
                        {
                            _originalXrefIds.Add(btrId);
                        }
                    }

                    trans.Commit();
                }

                ed.WriteMessage($"\nFound {_blockIdsBeforeBind.Count} total blocks and {_originalXrefIds.Count} DWG XREFs to bind.");

                if (_originalXrefIds.Count > 0)
                {
                    var idsToBind = new ObjectIdCollection(_originalXrefIds.ToArray());
                    if (!SkipBindDuringFinalize)
                    {
                        ed.WriteMessage($"\nBinding {idsToBind.Count} DWG reference(s)...");
                        db.BindXrefs(idsToBind, true);
                        bindCount = idsToBind.Count;
                    }
                    else
                    {
                        ed.WriteMessage("\nSkipping XREF binding per configuration; will detach originals only.");
                    }
                }

                if (bindCount > 0 || ForceDetachOriginalXrefs || _originalXrefIds.Count > 0)
                {
                    ed.WriteMessage("\nBind complete or skipped. Queueing cleanup process...");
                    doc.SendStringToExecute("_-FINALIZE-CLEANUP ", true, false, false);
                }
                else
                {
                    ed.WriteMessage("\nNo bindable DWG references found.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during bind: {ex.Message}");
            }
            finally
            {
                SkipBindDuringFinalize = false;
            }
        }
    }
}