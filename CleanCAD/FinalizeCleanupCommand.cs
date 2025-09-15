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
        // STAGE 2: Erase images, build the "kill list", and detach ghost XREF blocks.
        [CommandMethod("-FINALIZE-CLEANUP", CommandFlags.Modal)]
        public static void FinalizeCleanupCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;
        
            ed.WriteMessage("\n--- Stage 2: Erasing images and detaching ghost XREF blocks... ---");
        
            try
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    var newBlockIds = new List<ObjectId>();
                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    foreach (ObjectId currentBtrId in bt)
                    {
                        if (!_blockIdsBeforeBind.Contains(currentBtrId))
                        {
                            newBlockIds.Add(currentBtrId);
                        }
                    }
                    ed.WriteMessage($"\nFound {newBlockIds.Count} new block definition(s) created by bind.");
        
                    int imagesErased = 0;
                    foreach (ObjectId newBtrId in newBlockIds)
                    {
                        BlockTableRecord newBtr = trans.GetObject(newBtrId, OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId entId in newBtr)
                        {
                            if (entId.ObjectClass.DxfName == "IMAGE")
                            {
                                RasterImage image = trans.GetObject(entId, OpenMode.ForWrite) as RasterImage;
                                if (image != null && !image.ImageDefId.IsNull)
                                {
                                    _imageDefsToPurge.Add(image.ImageDefId); // Add the definition to our kill list
                                }
                                image.Erase();
                                imagesErased++;
                            }
                        }
                    }
                    ed.WriteMessage($"\nErased {imagesErased} RasterImage entit(ies) and added {_imageDefsToPurge.Count} definitions to kill list.");
        
                    int ghostsDetached = 0;
                    foreach (ObjectId originalXrefId in _originalXrefIds)
                    {
                        var obj = trans.GetObject(originalXrefId, OpenMode.ForRead, false, true);
                        if (obj != null && !obj.IsErased)
                        {
                            db.DetachXref(originalXrefId);
                            ghostsDetached++;
                        }
                    }
                    ed.WriteMessage($"\nManually detached {ghostsDetached} old XREF block definition(s).");
                    
                    trans.Commit();
                }
        
                ed.WriteMessage("\nIntermediate cleanup complete. Queueing final surgical purge...");
                doc.SendStringToExecute("_-FINALIZE-PURGEDEFS ", true, false, false);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during cleanup: {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                _blockIdsBeforeBind.Clear();
                _originalXrefIds.Clear();
            }
        }
    }
}

