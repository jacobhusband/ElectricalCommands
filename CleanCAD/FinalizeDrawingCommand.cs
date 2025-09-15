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
        // STAGE 1: Get snapshots, bind, and queue the main cleanup command.
        [CommandMethod("FINALIZE")]
        public static void FinalizeDrawingCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;
        
            _blockIdsBeforeBind.Clear();
            _originalXrefIds.Clear();
            _imageDefsToPurge.Clear();
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
                        if (btr != null && btr.IsFromExternalReference && btr.IsResolved &&
                            !string.IsNullOrEmpty(btr.PathName) && btr.PathName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
                        {
                            _originalXrefIds.Add(btrId);
                        }
                    }
                    ed.WriteMessage($"\nFound {_blockIdsBeforeBind.Count} total blocks and {_originalXrefIds.Count} DWG XREFs to bind.");
        
                    if (_originalXrefIds.Count > 0)
                    {
                        var idsToBind = new ObjectIdCollection(_originalXrefIds.ToArray());
                        ed.WriteMessage($"\nBinding {idsToBind.Count} DWG reference(s)...");
                        db.BindXrefs(idsToBind, true);
                        bindCount = idsToBind.Count;
                    }
                    trans.Commit();
                }
        
                if (bindCount > 0)
                {
                    ed.WriteMessage("\nBind complete. Queueing cleanup process...");
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
        }
    }
}

