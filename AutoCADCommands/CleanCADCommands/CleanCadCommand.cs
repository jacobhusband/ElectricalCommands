using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        private enum CleanWorkflowKind
        {
            TitleBlock,
            Sheet
        }

        [CommandMethod("CLEANTBLK", CommandFlags.Modal)]
        public static void RunCleanTitleBlock()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // New: Explode the main title block reference if it exists, to expose nested images.
            FindAndExplodeTitleBlockReference();

            // Leave the title-block workflow as-is.
            RunCleanWorkflow(CleanWorkflowKind.TitleBlock);
        }

        [CommandMethod("CLEANSHEET", CommandFlags.Modal)]
        public static void RunCleanSheet()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                if (TryGetTitleBlockOutlinePointsForEmbed(db, out var tbPoly) && tbPoly != null && tbPoly.Length > 0)
                {
                    ed.WriteMessage("\nTitle block found, zooming in...");
                    ZoomToTitleBlockForEmbed(ed, tbPoly);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Warning] Could not zoom to title block: {ex.Message}");
            }

            try
            {
                // Set necessary flags for the entire workflow before starting.
                CleanupCommands.SkipBindDuringFinalize = false;
                CleanupCommands.ForceDetachOriginalXrefs = false;
                CleanupCommands.RunKeepOnlyAfterFinalize = false;
                _chainFinalizeAfterEmbed = false; // Ensure EMBEDFROMXREFS doesn't chain a command on its own.

                ed.WriteMessage("\nCLEANSHEET: Queuing EMBEDFROMXREFS → CLEANPS → VP2PL → FINALIZE ...");

                // Queue the entire sequence of commands in a single string.
                // AutoCAD will execute them one by one after the current command scope ends.
                // The space at the end ensures the last command is executed.
                doc.SendStringToExecute("_.EMBEDFROMXREFS _.CLEANPS _.VP2PL _.FINALIZE ", true, false, false);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nCLEANSHEET failed to queue commands: {ex.Message}");
            }
        }

        // Legacy alias should map to the new CLEANSHEET behavior (not the shared workflow).
        [CommandMethod("CLEANCAD", CommandFlags.Modal)]
        public static void RunCleanCad()
        {
            RunCleanSheet();
        }

        // Existing shared workflow remains for CLEANTBLK only.
        private static void RunCleanWorkflow(CleanWorkflowKind kind)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                PrepareXrefLayersForCleanup(db, ed);

                CleanupCommands.SkipBindDuringFinalize = kind == CleanWorkflowKind.TitleBlock;
                CleanupCommands.ForceDetachOriginalXrefs = kind == CleanWorkflowKind.TitleBlock;
                CleanupCommands.RunKeepOnlyAfterFinalize = kind == CleanWorkflowKind.Sheet;

                if (kind == CleanWorkflowKind.TitleBlock)
                {
                    CleanupCommands.KeepOnlyTitleBlockInModelSpace();
                    DetachSpecialXrefs();
                    _chainFinalizeAfterEmbed = true;
                    EmbedFromXrefs();
                }
                else
                {
                    // No longer used by CLEANSHEET — left intact for compatibility.
                    CleanupCommands.ListPaperSpaceXrefs();
                    CleanupCommands.ViewportToPolyline_AllLayouts();
                    DetachSpecialXrefs();
                    _chainFinalizeAfterEmbed = true;
                    EmbedFromXrefs();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nClean workflow failed: {ex.Message}");
                _chainFinalizeAfterEmbed = false;
                CleanupCommands.RunKeepOnlyAfterFinalize = false;
                CleanupCommands.SkipBindDuringFinalize = false;
                CleanupCommands.ForceDetachOriginalXrefs = false;
            }
        }

        public static void PrepareXrefLayersForCleanup(Database db, Editor ed)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has("0"))
                    {
                        ed.WriteMessage("\nLayer '0' was not found; skipping XREF layer normalization.");
                        tr.Commit();
                        return;
                    }

                    var zeroId = lt["0"];
                    var zeroLayer = (LayerTableRecord)tr.GetObject(zeroId, OpenMode.ForWrite);
                    bool zeroWasLocked = zeroLayer.IsLocked;
                    if (zeroLayer.IsFrozen) zeroLayer.IsFrozen = false;
                    if (zeroLayer.IsOff) zeroLayer.IsOff = false;
                    if (zeroWasLocked) zeroLayer.IsLocked = false;

                    try
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        int moved = 0;
                        foreach (ObjectId recordId in bt)
                        {
                            var space = (BlockTableRecord)tr.GetObject(recordId, OpenMode.ForRead);
                            foreach (ObjectId entId in space)
                            {
                                var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                                if (br == null) continue;

                                BlockTableRecord def = null;
                                try { def = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord; }
                                catch { }

                                if (def == null || !def.IsFromExternalReference) continue;
                                if (br.LayerId == zeroId) continue;

                                LayerTableRecord sourceLayer = null;
                                bool relockSource = false;
                                try { sourceLayer = (LayerTableRecord)tr.GetObject(br.LayerId, OpenMode.ForRead); }
                                catch { }

                                if (sourceLayer != null && sourceLayer.IsLocked)
                                {
                                    sourceLayer.UpgradeOpen();
                                    sourceLayer.IsLocked = false;
                                    relockSource = true;
                                }

                                br.UpgradeOpen();
                                br.LayerId = zeroId;
                                moved++;

                                if (relockSource && sourceLayer != null)
                                {
                                    sourceLayer.IsLocked = true;
                                }
                            }
                        }

                        if (moved > 0)
                        {
                            ed.WriteMessage($"\nMoved {moved} XREF block reference(s) to layer '0'.");
                        }

                        if (zeroWasLocked)
                        {
                            zeroLayer.IsLocked = true;
                        }

                        tr.Commit();
                    }
                    finally
                    {
                        if (zeroWasLocked && !zeroLayer.IsLocked)
                        {
                            try { zeroLayer.IsLocked = true; } catch { }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to normalize XREF layers: {ex.Message}");
            }
        }
    }
}