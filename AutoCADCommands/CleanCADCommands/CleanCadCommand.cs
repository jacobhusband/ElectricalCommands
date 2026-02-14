using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;
using System.Linq;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        internal static Point3d[] _lastFoundTitleBlockPoly = null;

        [CommandMethod("CLEANTBLK", CommandFlags.Modal)]
        public static void RunCleanTitleBlock()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                CleanupCommands.ResetStrictTitleBlockProtection();
                _lastFoundTitleBlockPoly = null;

                EnsureAllLayersVisibleAndUnlocked(db, ed);

                DetachSpecialXrefs();

                _skipLayerFreezing = true;

                PreExplosionCleaner.CleanBeforeExplode();

                ExplodeAllBlockReferences();

                PrepareXrefLayersForCleanup(db, ed);

                CleanupCommands.SkipBindDuringFinalize = true;
                CleanupCommands.ForceDetachOriginalXrefs = true;
                CleanupCommands.RunKeepOnlyAfterFinalize = false;
                CleanupCommands.RunRemoveRemainingAfterFinalize = true;

                CleanupCommands.KeepOnlyTitleBlockInModelSpace();

                if (_lastFoundTitleBlockPoly != null && _lastFoundTitleBlockPoly.Length > 0)
                {
                    ed.WriteMessage("\nZooming to selected title block region...");
                    CleanupCommands.ZoomToTitleBlock(ed, _lastFoundTitleBlockPoly);
                }

                DetachSpecialXrefs();
                _chainFinalizeAfterEmbed = true;
                doc.SendStringToExecute("_.EMBEDIMAGES ", true, false, false);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nClean workflow failed: {ex.Message}");
                // Reset flags on failure
                _chainFinalizeAfterEmbed = false;
                CleanupCommands.RunKeepOnlyAfterFinalize = false;
                CleanupCommands.SkipBindDuringFinalize = false;
                CleanupCommands.ForceDetachOriginalXrefs = false;
                CleanupCommands.RunRemoveRemainingAfterFinalize = false;
                _skipLayerFreezing = false;
            }
        }

        [CommandMethod("CLEANCAD", CommandFlags.Modal)]
        public static void RunCleanSheet()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            _lastFoundTitleBlockPoly = null; // Reset at the start of the workflow
            CleanupCommands.ResetStrictTitleBlockProtection();

            if (!TryConfigureTitleBlockProtectionForCleanCad(ed, db))
            {
                CleanupCommands.ResetStrictTitleBlockProtection();
                return;
            }

            try
            {
                if (TryGetTitleBlockOutlinePointsForEmbed(db, out var tbPoly) && tbPoly != null && tbPoly.Length > 0)
                {
                    _lastFoundTitleBlockPoly = tbPoly; // Keep the polygon for later
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
                _chainFinalizeAfterEmbed = false;
                _skipLayerFreezing = false; // Ensure freezing is NOT skipped for this workflow

                // *** MODIFICATION START ***
                // Set the workflow flag to true
                _isCleanSheetWorkflowActive = true;
                ed.WriteMessage("\nCLEANSHEET: Starting EMBEDIMAGES...");

                // Queue ONLY the FIRST command. The rest will be chained.
                doc.SendStringToExecute("_.EMBEDIMAGES ", true, false, false);
                // *** MODIFICATION END ***
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nCLEANSHEET failed to queue commands: {ex.Message}");
                _isCleanSheetWorkflowActive = false; // Reset flag on failure
                CleanupCommands.ResetStrictTitleBlockProtection();
            }
        }

        private static bool TryConfigureTitleBlockProtectionForCleanCad(Editor ed, Database db)
        {
            var resolution = TitleBlockXrefResolver.Resolve(db);
            if (resolution.Kind == TitleBlockResolutionKind.Resolved && resolution.Winner != null)
            {
                var winner = resolution.Winner;
                CleanupCommands.EnableStrictTitleBlockProtection(winner.XrefBtrId, winner.BlockName, winner.PathName, winner.LayoutName);

                string fileName = string.Empty;
                try { fileName = Path.GetFileName(winner.PathName ?? string.Empty) ?? string.Empty; } catch { }
                string fileSuffix = string.IsNullOrWhiteSpace(fileName) ? string.Empty : $" ({fileName})";

                ed.WriteMessage(
                    $"\nCLEANCAD: Protecting inferred titleblock XREF '{winner.BlockName}' on layout '{winner.LayoutName}'{fileSuffix}.");
                return true;
            }

            if (resolution.Candidates.Count > 0)
            {
                ed.WriteMessage("\nCLEANCAD: Titleblock inference was ambiguous. Top candidate XREFs:");
                foreach (var candidate in resolution.Candidates.Take(8))
                {
                    string fileName = string.Empty;
                    try { fileName = Path.GetFileName(candidate.PathName ?? string.Empty) ?? string.Empty; } catch { }
                    string fileSuffix = string.IsNullOrWhiteSpace(fileName) ? string.Empty : $" ({fileName})";
                    ed.WriteMessage(
                        $"\n - Score {candidate.Score}: Layout '{candidate.LayoutName}', XREF '{candidate.BlockName}'{fileSuffix}");
                }
            }
            else
            {
                ed.WriteMessage("\nCLEANCAD: No paper-space XREF candidates were inferred as a titleblock.");
            }

            if (!TryPromptForTitleBlockXrefSelection(
                ed,
                db,
                out ObjectId selectedXrefId,
                out string selectedName,
                out string selectedPath,
                out string selectedLayout))
            {
                ed.WriteMessage("\nCLEANCAD cancelled: a titleblock XREF selection is required.");
                return false;
            }

            CleanupCommands.EnableStrictTitleBlockProtection(selectedXrefId, selectedName, selectedPath, selectedLayout);
            ed.WriteMessage(
                $"\nCLEANCAD: Protecting user-selected titleblock XREF '{selectedName}' on layout '{selectedLayout}'.");
            return true;
        }

        private static bool TryPromptForTitleBlockXrefSelection(
            Editor ed,
            Database db,
            out ObjectId xrefId,
            out string xrefName,
            out string xrefPath,
            out string layoutName)
        {
            xrefId = ObjectId.Null;
            xrefName = string.Empty;
            xrefPath = string.Empty;
            layoutName = string.Empty;

            for (int attempt = 0; attempt < 3; attempt++)
            {
                var peo = new PromptEntityOptions("\nSelect the titleblock XREF block reference (ESC to cancel): ");
                peo.SetRejectMessage("\nSelection must be a block reference.");
                peo.AddAllowedClass(typeof(BlockReference), false);

                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    return false;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var br = tr.GetObject(per.ObjectId, OpenMode.ForRead, false) as BlockReference;
                    if (br == null)
                    {
                        ed.WriteMessage("\nSelection was not a block reference.");
                        tr.Commit();
                        continue;
                    }

                    var def = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead, false) as BlockTableRecord;
                    if (def == null || (!def.IsFromExternalReference && !def.IsFromOverlayReference))
                    {
                        ed.WriteMessage("\nSelected block is not an external reference.");
                        tr.Commit();
                        continue;
                    }

                    xrefId = br.BlockTableRecord;
                    xrefName = def.Name ?? string.Empty;
                    xrefPath = def.PathName ?? string.Empty;

                    var owner = tr.GetObject(br.OwnerId, OpenMode.ForRead, false) as BlockTableRecord;
                    if (owner != null && owner.IsLayout)
                    {
                        var layout = tr.GetObject(owner.LayoutId, OpenMode.ForRead, false) as Layout;
                        if (layout != null)
                        {
                            layoutName = layout.LayoutName ?? string.Empty;
                        }
                    }

                    tr.Commit();
                    return true;
                }
            }

            ed.WriteMessage("\nToo many invalid selections. Aborting.");
            return false;
        }

        [CommandMethod("ZOOMTOLASTTB", CommandFlags.Modal)]
        public static void ZoomToLastTitleBlock()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Use a transaction to ensure database operations are stable
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // Step 1: Ensure we are on a layout tab (not the Model tab).
                    // TILEMODE = 0 means a layout tab is active.
                    if (Convert.ToInt16(Application.GetSystemVariable("TILEMODE")) == 1)
                    {
                        ed.WriteMessage("\nSwitching from Model tab to a Layout tab...");
                        Application.SetSystemVariable("TILEMODE", 0);
                    }

                    // Step 2: Ensure we are in Paper Space, not a floating viewport.
                    // CVPORT = 1 refers to the main paper space viewport.
                    // Any other value means a floating model space viewport is active.
                    if (Convert.ToInt16(Application.GetSystemVariable("CVPORT")) != 1)
                    {
                        ed.WriteMessage("\nActivating Paper Space...");
                        ed.SwitchToPaperSpace();
                    }

                    // Optional: A final check to confirm the state.
                    if (Convert.ToInt16(Application.GetSystemVariable("CVPORT")) != 1)
                    {
                        ed.WriteMessage("\nWarning: Failed to switch to Paper Space. Aborting zoom.");
                        return; // Exit if we couldn't switch
                    }

                    tr.Commit(); // Commit transaction if all is well
                }
            }
            catch (System.Exception ex)
            {
                // This might fail if, for example, there are no layouts.
                ed.WriteMessage($"\nCould not activate Paper Space: {ex.Message}");
            }

            if (_lastFoundTitleBlockPoly != null && _lastFoundTitleBlockPoly.Length > 0)
            {
                ed.WriteMessage("\nZooming to the title block found at the start of the workflow...");
                // Use the public zoom helper from the CleanupCommands class
                CleanupCommands.ZoomToTitleBlock(ed, _lastFoundTitleBlockPoly);
            }
            else
            {
                ed.WriteMessage("\nNo title block geometry was stored to zoom to.");
            }
            _lastFoundTitleBlockPoly = null; // Clean up the static variable
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
