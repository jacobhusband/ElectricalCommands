using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        /// <summary>
        /// Initial Core Console-safe CLEANCAD workflow.
        ///
        /// This variant intentionally refuses drawings that still contain raster
        /// images or underlays. The desktop workflow currently converts that content
        /// to OLE through PowerPoint, the Windows clipboard, PASTECLIP, and dialog
        /// automation; none of those operations belong in an unattended worker.
        /// </summary>
        [CommandMethod("CLEANCAD2", CommandFlags.Modal)]
        public static void RunCleanSheetHeadless()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;

            ResetHeadlessWorkflowState();
            ed.WriteMessage("\n=== CLEANCAD2: Starting non-interactive cleanup ===");

            TitleBlockXrefResolutionResult resolution;
            try
            {
                resolution = TitleBlockXrefResolver.Resolve(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nCLEANCAD2_ERROR: Titleblock preflight failed: {ex.Message}");
                return;
            }

            if (resolution.Kind != TitleBlockResolutionKind.Resolved || resolution.Winner == null)
            {
                ReportHeadlessTitleBlockFailure(ed, resolution);
                return;
            }

            HeadlessPreflightResult preflight = InspectForHeadlessBlockers(db);
            if (!preflight.Succeeded)
            {
                ed.WriteMessage($"\nCLEANCAD2_ERROR: Drawing preflight failed: {preflight.FailureMessage}");
                return;
            }

            if (preflight.UnresolvedXrefs.Count > 0)
            {
                ed.WriteMessage(
                    $"\nCLEANCAD2_ERROR: {preflight.UnresolvedXrefs.Count} unresolved or unloaded DWG XREF(s) prevent a safe headless run:");
                foreach (string xref in preflight.UnresolvedXrefs.Take(12))
                {
                    ed.WriteMessage($"\n - {xref}");
                }
                ed.WriteMessage("\nCLEANCAD2: No drawing changes were made.");
                return;
            }

            if (preflight.RasterImageCount > 0 || preflight.UnderlayCount > 0)
            {
                ed.WriteMessage(
                    $"\nCLEANCAD2_ERROR: Headless media embedding is not implemented yet. " +
                    $"Found {preflight.RasterImageCount} raster image reference(s) and {preflight.UnderlayCount} underlay reference(s).");
                foreach (string location in preflight.MediaLocations.Take(12))
                {
                    ed.WriteMessage($"\n - {location}");
                }
                ed.WriteMessage(
                    "\nCLEANCAD2: Aborted before mutation so external media cannot be lost. Use CLEANCAD for this drawing until a headless embedding strategy is added.");
                return;
            }

            var winner = resolution.Winner;
            CleanupCommands.EnableStrictTitleBlockProtection(
                winner.XrefBtrId,
                winner.BlockName,
                winner.PathName,
                winner.LayoutName);

            string titleBlockFile = string.Empty;
            try { titleBlockFile = Path.GetFileName(winner.PathName ?? string.Empty) ?? string.Empty; }
            catch { }

            ed.WriteMessage(
                $"\nCLEANCAD2: Protecting inferred titleblock XREF '{winner.BlockName}' on layout '{winner.LayoutName}'" +
                (string.IsNullOrWhiteSpace(titleBlockFile) ? "." : $" ({titleBlockFile})."));

            CleanupCommands.SkipBindDuringFinalize = false;
            CleanupCommands.UseClassicBindDuringFinalize = true;
            CleanupCommands.ForceDetachOriginalXrefs = false;
            CleanupCommands.RunKeepOnlyAfterFinalize = false;
            CleanupCommands.RunRemoveRemainingAfterFinalize = false;
            CleanupCommands.EnableModelspaceXrefCopyFallback = false;
            CleanupCommands.RunFinalizeStagesSynchronously = true;
            CleanupCommands.PreserveRemainingXrefsDuringFinalizeCleanup = true;
            _skipLayerFreezing = false;

            try
            {
                ed.WriteMessage("\nCLEANCAD2: Exploding non-attributed local Model Space blocks...");
                ExplodeAllBlockReferencesSkippingAttributed();

                ed.WriteMessage("\nCLEANCAD2: Binding XREFs and running finalize stages synchronously...");
                CleanupCommands.FinalizeDrawingCommand();

                if (CleanupCommands.FinalizeStageFailed ||
                    CleanupCommands.AbortRemainingXrefDetach ||
                    CleanupCommands.StrictTitleBlockBindFailed)
                {
                    ed.WriteMessage(
                        "\nCLEANCAD2_ERROR: XREF finalization did not pass strict titleblock verification. Remaining cleanup was stopped.");
                    return;
                }

                List<string> remainingXrefs = FindRemainingDwgXrefs(db);
                if (remainingXrefs.Count > 0)
                {
                    ed.WriteMessage(
                        $"\nCLEANCAD2_ERROR: {remainingXrefs.Count} DWG XREF(s) remained external after binding. " +
                        "They were preserved and destructive cleanup was stopped:");
                    foreach (string xref in remainingXrefs.Take(12))
                    {
                        ed.WriteMessage($"\n - {xref}");
                    }
                    ed.WriteMessage("\nCLEANCAD2: Close this partially modified drawing without saving.");
                    return;
                }

                ed.WriteMessage("\nCLEANCAD2: Cleaning paper space without activating layouts...");
                CleanupCommands.CleanPaperSpaceHeadless();

                ed.WriteMessage("\nCLEANCAD2: Cleaning Model Space with database-only viewport intersection...");
                if (!CleanupCommands.RunViewportToPolylineHeadless())
                {
                    ed.WriteMessage(
                        "\nCLEANCAD2_ERROR: Model Space viewport cleanup could not be completed safely. Remaining XREF detachment was stopped.");
                    return;
                }

                ed.WriteMessage("\nCLEANCAD2: Removing remaining non-protected references...");
                RemoveRemainingXrefs();

                ed.WriteMessage(
                    "\nCLEANCAD2_COMPLETE: Non-interactive cleanup finished. Save the drawing from the calling script after checking this marker.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nCLEANCAD2_ERROR: Headless cleanup failed: {ex.Message}");
            }
            finally
            {
                CleanupCommands.RunFinalizeStagesSynchronously = false;
                CleanupCommands.SkipBindDuringFinalize = false;
                CleanupCommands.UseClassicBindDuringFinalize = false;
                CleanupCommands.ForceDetachOriginalXrefs = false;
                CleanupCommands.RunKeepOnlyAfterFinalize = false;
                CleanupCommands.RunRemoveRemainingAfterFinalize = false;
                CleanupCommands.EnableModelspaceXrefCopyFallback = false;
                CleanupCommands.PreserveRemainingXrefsDuringFinalizeCleanup = false;
                CleanupCommands.FinalizeStageFailed = false;
                _skipLayerFreezing = false;

                if (CleanupCommands.StrictTitleBlockProtectionActive)
                {
                    CleanupCommands.ResetStrictTitleBlockProtection();
                }
            }
        }

        private static void ResetHeadlessWorkflowState()
        {
            _lastFoundTitleBlockPoly = null;
            _chainFinalizeAfterEmbed = false;
            _isCleanSheetWorkflowActive = false;
            _isEmbeddingProcessActive = false;
            _skipLayerFreezing = false;
            CleanupCommands.RunFinalizeStagesSynchronously = false;
            CleanupCommands.RunKeepOnlyAfterFinalize = false;
            CleanupCommands.RunRemoveRemainingAfterFinalize = false;
            CleanupCommands.SkipBindDuringFinalize = false;
            CleanupCommands.ForceDetachOriginalXrefs = false;
            CleanupCommands.EnableModelspaceXrefCopyFallback = false;
            CleanupCommands.PreserveRemainingXrefsDuringFinalizeCleanup = false;
            CleanupCommands.FinalizeStageFailed = false;
            CleanupCommands.ResetStrictTitleBlockProtection();
        }

        private static void ReportHeadlessTitleBlockFailure(
            Editor ed,
            TitleBlockXrefResolutionResult resolution)
        {
            string reason = resolution == null || resolution.Kind == TitleBlockResolutionKind.NotFound
                ? "no titleblock XREF could be inferred"
                : "titleblock inference was ambiguous";

            ed.WriteMessage($"\nCLEANCAD2_ERROR: {reason}. Headless mode never prompts for an entity selection.");

            if (resolution?.Candidates != null)
            {
                foreach (TitleBlockXrefCandidate candidate in resolution.Candidates.Take(12))
                {
                    string fileName = string.Empty;
                    try { fileName = Path.GetFileName(candidate.PathName ?? string.Empty) ?? string.Empty; }
                    catch { }

                    ed.WriteMessage(
                        $"\n - Score {candidate.Score}: Layout '{candidate.LayoutName}', XREF '{candidate.BlockName}'" +
                        (string.IsNullOrWhiteSpace(fileName) ? string.Empty : $" ({fileName})"));
                }
            }

            ed.WriteMessage("\nCLEANCAD2: No drawing changes were made.");
        }

        private static HeadlessPreflightResult InspectForHeadlessBlockers(Database db)
        {
            var result = new HeadlessPreflightResult();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (blockTable == null)
                    {
                        result.FailureMessage = "Block table is unavailable.";
                        return result;
                    }

                    foreach (ObjectId blockId in blockTable)
                    {
                        if (!blockId.IsValid || blockId.IsErased) continue;

                        var block = tr.GetObject(blockId, OpenMode.ForRead, false) as BlockTableRecord;
                        if (block == null || block.IsErased) continue;

                        if ((block.IsFromExternalReference || block.IsFromOverlayReference) && !block.IsResolved)
                        {
                            string path = block.PathName ?? string.Empty;
                            result.UnresolvedXrefs.Add(
                                $"'{block.Name}' [{block.XrefStatus}]" +
                                (string.IsNullOrWhiteSpace(path) ? string.Empty : $" -> {path}"));
                        }

                        foreach (ObjectId entityId in block)
                        {
                            if (!entityId.IsValid || entityId.IsErased) continue;

                            Entity entity = null;
                            try
                            {
                                entity = tr.GetObject(entityId, OpenMode.ForRead, false) as Entity;
                            }
                            catch (System.Exception ex)
                            {
                                string handle = "<unknown>";
                                try { handle = entityId.Handle.ToString(); }
                                catch { }
                                result.FailureMessage =
                                    $"Entity {handle} in block/space '{block.Name}' could not be inspected: {ex.Message}";
                                return result;
                            }
                            if (entity == null || entity.IsErased) continue;

                            if (entity is RasterImage)
                            {
                                result.RasterImageCount++;
                                result.MediaLocations.Add(DescribeMediaLocation(block, entity, "RasterImage"));
                            }
                            else if (entity is UnderlayReference)
                            {
                                result.UnderlayCount++;
                                result.MediaLocations.Add(DescribeMediaLocation(block, entity, entity.GetType().Name));
                            }
                        }
                    }

                    tr.Commit();
                }

                result.Succeeded = true;
            }
            catch (System.Exception ex)
            {
                result.FailureMessage = ex.Message;
            }

            return result;
        }

        private static List<string> FindRemainingDwgXrefs(Database db)
        {
            var remaining = new List<string>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (blockTable == null)
                {
                    remaining.Add("<block table unavailable>");
                    return remaining;
                }

                foreach (ObjectId blockId in blockTable)
                {
                    if (!blockId.IsValid || blockId.IsErased) continue;
                    var block = tr.GetObject(blockId, OpenMode.ForRead, false, true) as BlockTableRecord;
                    if (block == null || block.IsErased) continue;
                    if (!block.IsFromExternalReference && !block.IsFromOverlayReference) continue;

                    remaining.Add(
                        $"'{block.Name}' [{block.XrefStatus}]" +
                        (string.IsNullOrWhiteSpace(block.PathName) ? string.Empty : $" -> {block.PathName}"));
                }

                tr.Commit();
            }

            return remaining;
        }

        private static string DescribeMediaLocation(BlockTableRecord block, Entity entity, string mediaType)
        {
            string owner = block?.Name ?? "<unknown block>";
            string handle = "<unknown>";
            try { handle = entity.Handle.ToString(); }
            catch { }
            return $"{mediaType} handle {handle} in block/space '{owner}'";
        }

        private sealed class HeadlessPreflightResult
        {
            internal bool Succeeded { get; set; }
            internal string FailureMessage { get; set; } = string.Empty;
            internal int RasterImageCount { get; set; }
            internal int UnderlayCount { get; set; }
            internal List<string> UnresolvedXrefs { get; } = new List<string>();
            internal List<string> MediaLocations { get; } = new List<string>();
        }
    }
}
