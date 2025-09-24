using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

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

            LogTitleBlockReferenceStatus(db, ed, "CLEANUP - Start");

            ed.WriteMessage("\n--- Stage 2: Cleaning up bound blocks and detaching remaining XREFs... ---");

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
                                    _imageDefsToPurge.Add(image.ImageDefId);
                                }
                                image.Erase();
                                imagesErased++;
                            }
                        }
                    }
                    ed.WriteMessage($"\nErased {imagesErased} RasterImage entit(ies) and added {_imageDefsToPurge.Count} definitions to kill list.");

                    int ghostsDetached = 0;
                    ed.WriteMessage($"\n--- Scanning for remaining XREFs to detach ---");

                    var xrefsToDetach = new List<ObjectId>();
                    var blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    string[] titleBlockHints = { "x-tb", "title", "tblock", "border", "sheet" };

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = trans.GetObject(btrId, OpenMode.ForRead, false, true) as BlockTableRecord;
                        if (btr == null || btr.IsErased) continue;

                        if (btr.IsFromExternalReference)
                        {
                            ed.WriteMessage($"\n[LOG] Checking if XREF '{btr.Name}' should be detached...");
                            string btrNameLower = (btr.Name ?? string.Empty).ToLowerInvariant();
                            string path = (btr.PathName ?? string.Empty);
                            string fileNoExt = string.Empty;
                            try { fileNoExt = Path.GetFileNameWithoutExtension(path)?.ToLowerInvariant() ?? string.Empty; } catch { }

                            bool nameMatch = titleBlockHints.Any(h => btrNameLower.Contains(h));
                            bool pathMatch = !string.IsNullOrEmpty(fileNoExt) && titleBlockHints.Any(h => fileNoExt.Contains(h));
                            bool isTitleBlock = nameMatch || pathMatch;

                            ed.WriteMessage($"[LOG]   - Name: '{btr.Name}', Path: '{path}'");
                            ed.WriteMessage($"[LOG]   - Name check match: {nameMatch}");
                            ed.WriteMessage($"[LOG]   - Path check match: {pathMatch}");
                            ed.WriteMessage($"[LOG]   - Is Title Block?: {isTitleBlock}");

                            if (isTitleBlock)
                            {
                                ed.WriteMessage($"\n  - RESULT: SKIPPING detachment of XREF '{btr.Name}'.");
                                continue;
                            }
                            else
                            {
                                ed.WriteMessage($"\n  - RESULT: QUEUING leftover XREF '{btr.Name}' for detachment.");
                                xrefsToDetach.Add(btrId);
                            }
                        }
                    }

                    foreach (var xrefId in xrefsToDetach)
                    {
                        try
                        {
                            var btrToDetach = trans.GetObject(xrefId, OpenMode.ForRead, false, true) as BlockTableRecord;
                            if (btrToDetach != null)
                            {
                                ed.WriteMessage($"\n[LOG] CALLING db.DetachXref() on '{btrToDetach.Name}' (Handle: {btrToDetach.Handle})...");
                                db.DetachXref(xrefId);
                                ed.WriteMessage($"\n[LOG] RETURNED from db.DetachXref() on '{btrToDetach.Name}'.");
                                ghostsDetached++;
                                LogTitleBlockReferenceStatus(db, ed, $"CLEANUP - After Detaching '{btrToDetach.Name}'");
                            }
                        }
                        catch (System.Exception exDetach)
                        {
                            ed.WriteMessage($"\n  - Failed to detach XREF {xrefId}: {exDetach.Message}");
                            LogTitleBlockReferenceStatus(db, ed, $"CLEANUP - After Detach FAILED for {xrefId}");
                        }
                    }
                    ed.WriteMessage($"\n--- End of XREF detachment ---");
                    ed.WriteMessage($"\nDetached {ghostsDetached} remaining XREF definition(s).");

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
                LogTitleBlockReferenceStatus(db, ed, "CLEANUP - End");
                _blockIdsBeforeBind.Clear();
                _originalXrefIds.Clear();
                ForceDetachOriginalXrefs = false;
            }
        }
    }
}