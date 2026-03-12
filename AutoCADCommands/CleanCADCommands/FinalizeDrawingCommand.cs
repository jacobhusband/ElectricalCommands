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
            bool useClassicBind = UseClassicBindDuringFinalize;
            bool insertBind = !useClassicBind;
            ObjectId protectedXrefId = ProtectedTitleBlockXrefId;
            StrictTitleBlockBindFailed = false;
            AbortRemainingXrefDetach = false;
            string protectedNameAtScan = ProtectedTitleBlockName;
            string protectedPathAtScan = ProtectedTitleBlockPath;

            _blockIdsBeforeBind.Clear();
            _originalXrefIds.Clear();
            _imageDefsToPurge.Clear();

            SimplerCommands.DetachSpecialXrefs();

            ed.WriteMessage("\n--- Stage 1: Analyzing and Binding... ---");
            if (useClassicBind)
            {
                ed.WriteMessage("\nCLEANCAD color-preserve mode: using classic bind so bound layers keep XREF prefixes such as 'xref$0$Layer'.");
            }

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
                bool protectedDwgRefreshActive = false;

                if (strictProtection)
                {
                    if (!TryGetXrefState(db, protectedXrefId, out XrefState stateBeforeRefresh) ||
                        !stateBeforeRefresh.Exists ||
                        !stateBeforeRefresh.IsExternal)
                    {
                        string statusDetail = stateBeforeRefresh.Exists
                            ? stateBeforeRefresh.StatusText
                            : "Missing";
                        string rawPathDetail = stateBeforeRefresh.Exists
                            ? stateBeforeRefresh.PathName
                            : protectedPathAtScan;
                        MarkStrictBindFailureAndAbortDetach(
                            ed,
                            db,
                            "protected titleblock is not a DWG XREF",
                            protectedNameAtScan,
                            rawPathDetail,
                            statusDetail);
                        return;
                    }

                    protectedNameAtScan = stateBeforeRefresh.Name ?? protectedNameAtScan;
                    protectedPathAtScan = stateBeforeRefresh.PathName ?? protectedPathAtScan;

                    if (!stateBeforeRefresh.IsDwg)
                    {
                        ed.WriteMessage(
                            $"\nCLEANCAD strict mode: protected titleblock '{protectedNameAtScan}' is not a DWG XREF. Skipping the pre-bind refresh/rebind path for this reference.");
                    }
                    else
                    {
                        protectedDwgRefreshActive = true;

                        if (!stateBeforeRefresh.IsResolved)
                        {
                            ed.WriteMessage(
                                $"\nProtected titleblock '{protectedNameAtScan}' is unresolved. Attempting one XREF reload...");
                            TryReloadXrefDefinition(db, ed, protectedXrefId);
                        }

                        string refreshFailure;
                        if (!TryRefreshProtectedTitleBlockForBind(
                            db,
                            ed,
                            ref protectedXrefId,
                            ref protectedNameAtScan,
                            ref protectedPathAtScan,
                            ProtectedTitleBlockLayoutName,
                            out refreshFailure))
                        {
                            MarkStrictBindFailureAndAbortDetach(
                                ed,
                                db,
                                $"failed to refresh protected titleblock before bind ({refreshFailure})",
                                protectedNameAtScan,
                                protectedPathAtScan,
                                "RefreshFailed");
                            return;
                        }

                        if (!TryGetXrefState(db, protectedXrefId, out XrefState refreshedState) ||
                            !refreshedState.Exists ||
                            !refreshedState.IsExternal ||
                            !refreshedState.IsResolved ||
                            !refreshedState.IsDwg ||
                            refreshedState.IsOverlay)
                        {
                            MarkStrictBindFailureAndAbortDetach(
                                ed,
                                db,
                                "protected titleblock did not become an attached, bindable DWG XREF after refresh",
                                protectedNameAtScan,
                                protectedPathAtScan,
                                refreshedState.StatusText);
                            return;
                        }

                        protectedNameAtScan = refreshedState.Name;
                        protectedPathAtScan = refreshedState.PathName;
                        EnableStrictTitleBlockProtection(
                            protectedXrefId,
                            refreshedState.Name,
                            refreshedState.PathName,
                            ProtectedTitleBlockLayoutName);
                        ed.WriteMessage(
                            $"\nCLEANCAD strict mode: titleblock '{refreshedState.Name}' was refreshed and queued for bind.");
                    }
                }

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

                ed.WriteMessage($"\nFound {_blockIdsBeforeBind.Count} total blocks and {_originalXrefIds.Count} DWG XREFs to bind.");

                if (_originalXrefIds.Count > 0)
                {
                    var idsToBind = new ObjectIdCollection(_originalXrefIds.ToArray());
                    if (!SkipBindDuringFinalize)
                    {
                        if (useClassicBind)
                        {
                            ed.WriteMessage($"\nBinding {idsToBind.Count} DWG reference(s) using classic bind to preserve XREF layer colors...");
                        }
                        else
                        {
                            ed.WriteMessage($"\nBinding {idsToBind.Count} DWG reference(s)...");
                        }

                        db.BindXrefs(idsToBind, insertBind);
                        bindCount = idsToBind.Count;
                    }
                    else
                    {
                        ed.WriteMessage("\nSkipping XREF binding per configuration; will detach originals only.");
                    }
                }

                if (protectedDwgRefreshActive && !SkipBindDuringFinalize)
                {
                    if (TryGetXrefState(db, protectedXrefId, out XrefState postBindState) &&
                        postBindState.Exists &&
                        postBindState.IsExternal)
                    {
                        MarkStrictBindFailureAndAbortDetach(
                            ed,
                            db,
                            "titleblock XREF remained external after bind",
                            postBindState.Name,
                            postBindState.PathName,
                            postBindState.StatusText);
                        return;
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
                UseClassicBindDuringFinalize = false;
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

        private static bool TryRefreshProtectedTitleBlockForBind(
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
                ObjectIdCollection refIds = null;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var oldBtr = tr.GetObject(protectedXrefId, OpenMode.ForRead, false, true) as BlockTableRecord;
                    if (oldBtr == null || oldBtr.IsErased)
                    {
                        failureReason = "protected xref definition is missing";
                        tr.Commit();
                        return false;
                    }

                    if (!oldBtr.IsFromExternalReference && !oldBtr.IsFromOverlayReference)
                    {
                        failureReason = "protected xref is no longer external";
                        tr.Commit();
                        return false;
                    }

                    if (!IsDwgXrefDefinition(oldBtr))
                    {
                        failureReason = "protected xref is not a DWG reference";
                        tr.Commit();
                        return false;
                    }

                    sourceName = oldBtr.Name ?? sourceName;
                    sourcePath = oldBtr.PathName ?? sourcePath;
                    refIds = oldBtr.GetBlockReferenceIds(true, true);
                    if (refIds == null || refIds.Count == 0)
                    {
                        failureReason = "protected xref has no block references to preserve";
                        tr.Commit();
                        return false;
                    }
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

                using (var tr = db.TransactionManager.StartTransaction())
                {
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
                    failureReason = $"failed to detach original titleblock xref ({detachEx.Message})";
                    return false;
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
                    $"\nRefreshed protected titleblock XREF '{sourceName}' with a fresh attached definition '{protectedName}' using '{protectedPath}'. Repointed {retargetedRefs} reference(s).");

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
