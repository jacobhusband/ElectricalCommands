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
            bool strictProtection = StrictTitleBlockProtectionActive;
            ObjectId protectedXrefId = ProtectedTitleBlockXrefId;
            StrictTitleBlockBindFailed = false;
            AbortRemainingXrefDetach = false;
            bool protectedPromotedToAttach = false;

            _blockIdsBeforeBind.Clear();
            _originalXrefIds.Clear();
            _imageDefsToPurge.Clear();

            SimplerCommands.DetachSpecialXrefs();

            ed.WriteMessage("\n--- Stage 1: Analyzing and Binding... ---");

            try
            {
                if (strictProtection && protectedXrefId.IsNull)
                {
                    MarkStrictBindFailureAndAbortDetach(
                        ed,
                        db,
                        "protected titleblock XREF id was not initialized",
                        ProtectedTitleBlockName,
                        ProtectedTitleBlockPath,
                        "MissingId");
                    return;
                }

                int bindCount = 0;
                bool protectedFound = false;
                bool protectedResolved = false;
                bool protectedIsDwg = false;
                string protectedNameAtScan = ProtectedTitleBlockName;
                string protectedPathAtScan = ProtectedTitleBlockPath;

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    foreach (ObjectId btrId in bt)
                    {
                        _blockIdsBeforeBind.Add(btrId);
                        var btr = trans.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr == null)
                        {
                            continue;
                        }

                        if (strictProtection && btrId == protectedXrefId)
                        {
                            protectedFound = true;
                            protectedResolved = btr.IsResolved;
                            protectedIsDwg = IsDwgXrefDefinition(btr);
                            protectedNameAtScan = btr.Name ?? protectedNameAtScan;
                            protectedPathAtScan = btr.PathName ?? protectedPathAtScan;
                        }

                        if (!btr.IsFromExternalReference || !btr.IsResolved)
                        {
                            continue;
                        }

                        if (IsDwgXrefDefinition(btr))
                        {
                            _originalXrefIds.Add(btrId);
                        }
                    }

                    trans.Commit();
                }

                if (strictProtection)
                {
                    if (!protectedFound)
                    {
                        MarkStrictBindFailureAndAbortDetach(
                            ed,
                            db,
                            "the protected titleblock XREF no longer exists in this drawing",
                            protectedNameAtScan,
                            protectedPathAtScan,
                            "Missing");
                        return;
                    }

                    if (!protectedIsDwg)
                    {
                        MarkStrictBindFailureAndAbortDetach(
                            ed,
                            db,
                            "protected titleblock is not a DWG XREF",
                            protectedNameAtScan,
                            protectedPathAtScan,
                            "NotDwg");
                        return;
                    }

                    if (!protectedResolved)
                    {
                        ed.WriteMessage(
                            $"\nProtected titleblock '{protectedNameAtScan}' is unresolved. Attempting one XREF reload...");
                        TryReloadXrefDefinition(db, ed, protectedXrefId);
                    }

                    if (!TryGetXrefState(db, protectedXrefId, out XrefState stateAfterReload) ||
                        !stateAfterReload.Exists ||
                        !stateAfterReload.IsExternal ||
                        !stateAfterReload.IsResolved ||
                        !stateAfterReload.IsDwg)
                    {
                        string statusDetail = stateAfterReload.Exists
                            ? stateAfterReload.StatusText
                            : "Missing";
                        string rawPathDetail = stateAfterReload.Exists
                            ? stateAfterReload.PathName
                            : protectedPathAtScan;
                        MarkStrictBindFailureAndAbortDetach(
                            ed,
                            db,
                            "titleblock is not bindable",
                            protectedNameAtScan,
                            rawPathDetail,
                            statusDetail);
                        return;
                    }

                    if (stateAfterReload.IsOverlay)
                    {
                        ed.WriteMessage(
                            $"\nProtected titleblock '{stateAfterReload.Name}' is an overlay XREF. Promoting to attached before bind...");

                        string promoteFailure;
                        if (!TryPromoteProtectedTitleBlockToAttach(
                            db,
                            ed,
                            ref protectedXrefId,
                            ref protectedNameAtScan,
                            ref protectedPathAtScan,
                            ProtectedTitleBlockLayoutName,
                            out promoteFailure))
                        {
                            MarkStrictBindFailureAndAbortDetach(
                                ed,
                                db,
                                $"failed to promote protected titleblock to attached ({promoteFailure})",
                                protectedNameAtScan,
                                protectedPathAtScan,
                                "PromotionFailed");
                            return;
                        }

                        protectedPromotedToAttach = true;

                        if (!TryGetXrefState(db, protectedXrefId, out stateAfterReload) ||
                            !stateAfterReload.Exists ||
                            !stateAfterReload.IsExternal ||
                            !stateAfterReload.IsResolved ||
                            !stateAfterReload.IsDwg ||
                            stateAfterReload.IsOverlay)
                        {
                            MarkStrictBindFailureAndAbortDetach(
                                ed,
                                db,
                                "protected titleblock did not become an attached, bindable DWG XREF after promotion",
                                protectedNameAtScan,
                                protectedPathAtScan,
                                stateAfterReload.StatusText);
                            return;
                        }
                    }

                    _originalXrefIds.Add(protectedXrefId);
                    EnableStrictTitleBlockProtection(
                        protectedXrefId,
                        stateAfterReload.Name,
                        stateAfterReload.PathName,
                        ProtectedTitleBlockLayoutName);
                    ed.WriteMessage(
                        $"\nCLEANCAD strict mode: titleblock '{stateAfterReload.Name}' is resolved and queued for bind.");
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

                if (strictProtection && !SkipBindDuringFinalize)
                {
                    if (TryGetXrefState(db, protectedXrefId, out XrefState postBindState) &&
                        postBindState.Exists &&
                        postBindState.IsExternal)
                    {
                        if (!protectedPromotedToAttach)
                        {
                            ed.WriteMessage(
                                $"\nProtected titleblock '{postBindState.Name}' remained external after bind. Retrying via attached promotion...");

                            string promoteFailure;
                            string postBindName = postBindState.Name;
                            string postBindPath = postBindState.PathName;

                            if (TryPromoteProtectedTitleBlockToAttach(
                                db,
                                ed,
                                ref protectedXrefId,
                                ref postBindName,
                                ref postBindPath,
                                ProtectedTitleBlockLayoutName,
                                out promoteFailure))
                            {
                                protectedPromotedToAttach = true;

                                _originalXrefIds.Add(protectedXrefId);
                                var retryIds = new ObjectIdCollection(new[] { protectedXrefId });
                                ed.WriteMessage("\nBinding promoted protected titleblock XREF...");
                                db.BindXrefs(retryIds, true);
                            }
                            else
                            {
                                MarkStrictBindFailureAndAbortDetach(
                                    ed,
                                    db,
                                    $"titleblock remained external after bind and promotion failed ({promoteFailure})",
                                    postBindState.Name,
                                    postBindState.PathName,
                                    postBindState.StatusText);
                                return;
                            }
                        }

                        if (TryGetXrefState(db, protectedXrefId, out XrefState postRetryState) &&
                            postRetryState.Exists &&
                            postRetryState.IsExternal)
                        {
                            MarkStrictBindFailureAndAbortDetach(
                                ed,
                                db,
                                "titleblock XREF remained external after bind",
                                postRetryState.Name,
                                postRetryState.PathName,
                                postRetryState.StatusText);
                            return;
                        }
                    }

                    ed.WriteMessage("\nCLEANCAD strict mode: titleblock bind verification passed.");
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
                if (strictProtection)
                {
                    MarkStrictBindFailureAndAbortDetach(
                        ed,
                        db,
                        "bind process threw an exception",
                        ProtectedTitleBlockName,
                        ProtectedTitleBlockPath,
                        ex.GetType().Name);
                }
            }
            finally
            {
                SkipBindDuringFinalize = false;
            }
        }

        private static void MarkStrictBindFailureAndAbortDetach(
            Editor ed,
            Database db,
            string failureDetail,
            string blockName,
            string rawPath,
            string statusText)
        {
            MarkStrictTitleBlockBindFailure();

            string safeBlock = string.IsNullOrWhiteSpace(blockName) ? "<unknown>" : blockName;
            string safeRawPath = string.IsNullOrWhiteSpace(rawPath) ? "<empty>" : rawPath;
            string resolvedPath = ResolveXrefPathForComparison(db, rawPath);
            string safeResolvedPath = string.IsNullOrWhiteSpace(resolvedPath) ? "<unresolved>" : resolvedPath;
            string safeStatus = string.IsNullOrWhiteSpace(statusText) ? "Unknown" : statusText;

            ed.WriteMessage(
                $"\nCLEANCAD strict mode failed: {failureDetail}. Titleblock: '{safeBlock}'. Status: {safeStatus}. Raw path: {safeRawPath}. Resolved path: {safeResolvedPath}.");
            ed.WriteMessage("\nCLEANCAD strict mode: downstream XREF detachment is now aborted for this run.");
        }

        private static bool TryPromoteProtectedTitleBlockToAttach(
            Database db,
            Editor ed,
            ref ObjectId protectedXrefId,
            ref string protectedName,
            ref string protectedPath,
            string protectedLayoutName,
            out string failureReason)
        {
            failureReason = string.Empty;

            if (db == null || protectedXrefId.IsNull)
            {
                failureReason = "missing protected xref id";
                return false;
            }

            string sourceName = protectedName ?? string.Empty;
            string sourcePath = protectedPath ?? string.Empty;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var oldBtr = tr.GetObject(protectedXrefId, OpenMode.ForRead, false, true) as BlockTableRecord;
                    if (oldBtr == null || oldBtr.IsErased)
                    {
                        failureReason = "protected xref definition is missing";
                        tr.Commit();
                        return false;
                    }

                    sourceName = oldBtr.Name ?? sourceName;
                    sourcePath = oldBtr.PathName ?? sourcePath;
                    tr.Commit();
                }

                string attachPath = ResolveXrefPathForComparison(db, sourcePath);
                if (string.IsNullOrWhiteSpace(attachPath))
                {
                    attachPath = NormalizeXrefPathToken(sourcePath);
                }

                if (string.IsNullOrWhiteSpace(attachPath))
                {
                    failureReason = "no usable source path";
                    return false;
                }

                string uniqueName = GetUniqueXrefBlockName(db, sourceName);
                ObjectId newXrefId = db.AttachXref(attachPath, uniqueName);
                if (newXrefId.IsNull)
                {
                    failureReason = "AttachXref returned null id";
                    return false;
                }

                int retargetedRefs = 0;
                var unlockedLayers = new HashSet<ObjectId>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var oldBtr = tr.GetObject(protectedXrefId, OpenMode.ForRead, false, true) as BlockTableRecord;
                    if (oldBtr == null || oldBtr.IsErased)
                    {
                        failureReason = "original xref disappeared during promotion";
                        tr.Commit();
                        return false;
                    }

                    ObjectIdCollection refIds = oldBtr.GetBlockReferenceIds(true, true);
                    foreach (ObjectId refId in refIds)
                    {
                        var br = tr.GetObject(refId, OpenMode.ForWrite, false, true) as BlockReference;
                        if (br == null || br.IsErased) continue;

                        var layer = tr.GetObject(br.LayerId, OpenMode.ForRead, false) as LayerTableRecord;
                        bool relock = false;
                        if (layer != null && layer.IsLocked)
                        {
                            layer.UpgradeOpen();
                            layer.IsLocked = false;
                            relock = true;
                            unlockedLayers.Add(br.LayerId);
                        }

                        br.BlockTableRecord = newXrefId;
                        retargetedRefs++;

                        if (relock && layer != null)
                        {
                            layer.IsLocked = true;
                        }
                    }

                    tr.Commit();
                }

                try
                {
                    db.DetachXref(protectedXrefId);
                }
                catch (System.Exception detachEx)
                {
                    ed.WriteMessage($"\nWarning: promoted titleblock old definition could not be detached: {detachEx.Message}");
                }

                if (!TryGetXrefState(db, newXrefId, out XrefState promotedState) ||
                    !promotedState.Exists ||
                    !promotedState.IsExternal ||
                    !promotedState.IsResolved ||
                    !promotedState.IsDwg)
                {
                    failureReason = "promoted xref state invalid";
                    return false;
                }

                protectedXrefId = newXrefId;
                protectedName = promotedState.Name;
                protectedPath = promotedState.PathName;

                EnableStrictTitleBlockProtection(
                    protectedXrefId,
                    protectedName,
                    protectedPath,
                    protectedLayoutName);

                ed.WriteMessage(
                    $"\nPromoted protected titleblock XREF to attached definition '{protectedName}' using '{protectedPath}'. Repointed {retargetedRefs} reference(s).");

                return true;
            }
            catch (System.Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
        }

        private static string GetUniqueXrefBlockName(Database db, string preferred)
        {
            string baseName = string.IsNullOrWhiteSpace(preferred)
                ? "ProtectedTitleBlock"
                : preferred.Trim();
            string candidate = baseName;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (bt == null)
                {
                    tr.Commit();
                    return candidate;
                }

                if (!bt.Has(candidate))
                {
                    tr.Commit();
                    return candidate;
                }

                string suffixBase = baseName + "_TB_ATTACHED";
                if (!bt.Has(suffixBase))
                {
                    tr.Commit();
                    return suffixBase;
                }

                int i = 2;
                while (bt.Has(suffixBase + "_" + i))
                {
                    i++;
                }

                tr.Commit();
                return suffixBase + "_" + i;
            }
        }

        private static bool IsDwgXrefDefinition(BlockTableRecord btr)
        {
            if (btr == null) return false;

            string pathName = btr.PathName ?? string.Empty;
            if (!string.IsNullOrEmpty(pathName))
            {
                try
                {
                    if (string.Equals(Path.GetExtension(pathName), ".dwg", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                catch { }
            }

            string name = btr.Name ?? string.Empty;
            return name.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase);
        }

        private static void TryReloadXrefDefinition(Database db, Editor ed, ObjectId xrefId)
        {
            if (db == null || xrefId.IsNull) return;

            try
            {
                db.ReloadXrefs(new ObjectIdCollection(new[] { xrefId }));
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nUnable to reload protected titleblock XREF: {ex.Message}");
            }
        }

        private static bool TryGetXrefState(Database db, ObjectId xrefId, out XrefState state)
        {
            state = XrefState.Missing;
            if (db == null || xrefId.IsNull) return false;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    if (!xrefId.IsValid || xrefId.IsErased)
                    {
                        tr.Commit();
                        state = XrefState.Missing;
                        return false;
                    }

                    var btr = tr.GetObject(xrefId, OpenMode.ForRead, false, true) as BlockTableRecord;
                    if (btr == null || btr.IsErased)
                    {
                        tr.Commit();
                        state = XrefState.Missing;
                        return false;
                    }

                    state = new XrefState
                    {
                        Exists = true,
                        IsExternal = btr.IsFromExternalReference || btr.IsFromOverlayReference,
                        IsOverlay = btr.IsFromOverlayReference,
                        IsResolved = btr.IsResolved,
                        IsDwg = IsDwgXrefDefinition(btr),
                        Name = btr.Name ?? string.Empty,
                        PathName = btr.PathName ?? string.Empty,
                        StatusText = btr.XrefStatus.ToString()
                    };

                    tr.Commit();
                    return true;
                }
            }
            catch
            {
                state = XrefState.Missing;
                return false;
            }
        }

        private struct XrefState
        {
            internal bool Exists;
            internal bool IsExternal;
            internal bool IsOverlay;
            internal bool IsResolved;
            internal bool IsDwg;
            internal string Name;
            internal string PathName;
            internal string StatusText;

            internal static XrefState Missing => new XrefState
            {
                Exists = false,
                IsExternal = false,
                IsOverlay = false,
                IsResolved = false,
                IsDwg = false,
                Name = string.Empty,
                PathName = string.Empty,
                StatusText = "Missing"
            };
        }
    }
}
