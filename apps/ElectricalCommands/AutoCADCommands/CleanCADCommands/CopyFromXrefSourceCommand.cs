using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        // Per-XREF copy-from-source fallback. Runs after db.BindXrefs() in FINALIZE
        // for DWG XREFs that remained external. Opens the source DWG as a side db,
        // WblockClones its modelspace into a new local BTR, replaces each modelspace
        // BlockReference, and detaches the def if no paperspace refs remain. The
        // protected paperspace titleblock and any non-modelspace ref is never touched.

        internal sealed class XrefEmbedFailure
        {
            internal string XrefName = string.Empty;
            internal string SourcePath = string.Empty;
            internal string Reason = string.Empty;
            internal string Category = string.Empty;
        }

        private sealed class XrefEmbedResult
        {
            internal ObjectId XrefDefId;
            internal string XrefName = string.Empty;
            internal string SourcePath = string.Empty;
            internal int ModelspaceRefsReplaced;
            internal int PaperspaceRefsLeftAlone;
            internal ObjectId NewLocalBtrId = ObjectId.Null;
            internal bool Detached;
            internal bool Success;
            internal string FailureReason = string.Empty;
            internal string FailureCategory = string.Empty;
        }

        internal static bool TryEmbedFailedModelspaceXrefs(
            Database hostDb,
            IEnumerable<ObjectId> candidateXrefDefIds,
            out int embeddedCount,
            out int failedCount,
            out List<XrefEmbedFailure> failures)
        {
            embeddedCount = 0;
            failedCount = 0;
            failures = new List<XrefEmbedFailure>();

            if (hostDb == null || candidateXrefDefIds == null) return true;

            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            var inFlight = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (!string.IsNullOrWhiteSpace(hostDb.Filename))
                {
                    inFlight.Add(Path.GetFullPath(hostDb.Filename));
                }
            }
            catch { }

            foreach (var xrefId in candidateXrefDefIds)
            {
                if (xrefId.IsNull) continue;
                if (IsProtectedTitleBlockXref(xrefId)) continue;

                var result = EmbedSingleXref(hostDb, ed, xrefId, inFlight);

                if (result.Success)
                {
                    embeddedCount++;
                    ed?.WriteMessage(
                        $"\n    OK: {result.ModelspaceRefsReplaced} ref(s) replaced; " +
                        $"{result.PaperspaceRefsLeftAlone} ps ref(s) left alone; " +
                        $"detach={result.Detached}.");
                }
                else
                {
                    failedCount++;
                    ed?.WriteMessage($"\n    FAIL ({result.FailureCategory}): {result.FailureReason}");
                    failures.Add(new XrefEmbedFailure
                    {
                        XrefName = result.XrefName,
                        SourcePath = result.SourcePath,
                        Reason = result.FailureReason,
                        Category = result.FailureCategory
                    });
                }
            }

            return failedCount == 0;
        }

        private static XrefEmbedResult EmbedSingleXref(
            Database hostDb, Editor ed, ObjectId xrefDefId,
            HashSet<string> inFlightSourcePaths)
        {
            var result = new XrefEmbedResult { XrefDefId = xrefDefId };

            // Defense-in-depth: never embed the protected titleblock (cross-db ids never collide).
            if (IsProtectedTitleBlockXref(xrefDefId))
            {
                result.Success = true;
                result.FailureCategory = "Skipped";
                result.FailureReason = "Protected titleblock — skipped.";
                return result;
            }

            string xrefName, absolutePath, failureReason;
            if (!TryResolveXrefSourcePath(hostDb, xrefDefId, out xrefName, out absolutePath, out failureReason))
            {
                result.XrefName = xrefName;
                result.FailureCategory = "MissingFile";
                result.FailureReason = failureReason;
                return result;
            }
            result.XrefName = xrefName;
            result.SourcePath = absolutePath;

            ed?.WriteMessage($"\n  - Embedding '{xrefName}' from '{absolutePath}'...");

            if (inFlightSourcePaths.Contains(absolutePath))
            {
                result.FailureCategory = "Cycle";
                result.FailureReason = $"Circular xref reference: {absolutePath}";
                return result;
            }
            inFlightSourcePaths.Add(absolutePath);

            try
            {
                var sideDb = OpenSideDatabase(absolutePath, out failureReason);
                if (sideDb == null)
                {
                    result.FailureCategory = "ReadDwgFile";
                    result.FailureReason = failureReason;
                    return result;
                }

                using (sideDb)
                {
                    EmbedNestedXrefsInSideDb(sideDb, ed, inFlightSourcePaths);

                    string localBlockName = GetUniqueLocalBlockName(hostDb, xrefName);

                    int refsReplaced = 0;
                    int paperspaceRefsRemaining = 0;
                    ObjectId newBtrId = ObjectId.Null;

                    using (var hostTr = hostDb.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            var hostBt = (BlockTable)hostTr.GetObject(hostDb.BlockTableId, OpenMode.ForWrite);
                            var newBtr = new BlockTableRecord { Name = localBlockName };
                            newBtrId = hostBt.Add(newBtr);
                            hostTr.AddNewlyCreatedDBObject(newBtr, true);

                            int cloned = CloneModelspaceContents(hostDb, sideDb, newBtrId, out failureReason);
                            if (cloned < 0)
                            {
                                result.FailureCategory = "WblockClone";
                                result.FailureReason = failureReason;
                                hostTr.Abort();
                                return result;
                            }

                            refsReplaced = ReplaceModelspaceBlockReferences(
                                hostDb, xrefDefId, newBtrId, hostTr,
                                out paperspaceRefsRemaining, out failureReason);
                            if (refsReplaced < 0)
                            {
                                result.FailureCategory = "ReplaceRefs";
                                result.FailureReason = failureReason;
                                hostTr.Abort();
                                return result;
                            }

                            hostTr.Commit();
                            result.NewLocalBtrId = newBtrId;
                        }
                        catch (System.Exception ex)
                        {
                            hostTr.Abort();
                            result.FailureCategory = "ReplaceRefs";
                            result.FailureReason = ex.Message;
                            return result;
                        }
                    }

                    result.ModelspaceRefsReplaced = refsReplaced;
                    result.PaperspaceRefsLeftAlone = paperspaceRefsRemaining;

                    if (paperspaceRefsRemaining == 0)
                    {
                        if (TryDetachIfFullyReplaced(hostDb, xrefDefId, out failureReason))
                        {
                            result.Detached = true;
                        }
                        else
                        {
                            ed?.WriteMessage($"\n    WARN: detach failed: {failureReason}");
                        }
                    }

                    result.Success = true;
                    return result;
                }
            }
            finally
            {
                inFlightSourcePaths.Remove(absolutePath);
            }
        }

        private static bool TryResolveXrefSourcePath(
            Database hostDb, ObjectId xrefDefId,
            out string xrefName, out string absolutePath, out string failureReason)
        {
            xrefName = string.Empty;
            absolutePath = string.Empty;
            failureReason = string.Empty;

            try
            {
                string raw;
                using (var tr = hostDb.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(xrefDefId, OpenMode.ForRead) as BlockTableRecord;
                    if (btr == null)
                    {
                        tr.Commit();
                        failureReason = "XREF definition could not be opened.";
                        return false;
                    }
                    xrefName = btr.Name ?? string.Empty;
                    raw = btr.PathName ?? string.Empty;
                    tr.Commit();
                }

                if (string.IsNullOrWhiteSpace(raw))
                {
                    failureReason = "XREF has no path.";
                    return false;
                }

                string resolved = ResolveXrefPathForComparison(hostDb, raw);
                if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                {
                    failureReason = $"Source file not found: {raw}";
                    return false;
                }
                absolutePath = Path.GetFullPath(resolved);
                return true;
            }
            catch (System.Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
        }

        private static Database OpenSideDatabase(string absolutePath, out string failureReason)
        {
            failureReason = string.Empty;
            Database sideDb = null;
            try
            {
                sideDb = new Database(false, true);
                sideDb.ReadDwgFile(absolutePath, FileShare.Read, true, null);
                sideDb.CloseInput(true);
                return sideDb;
            }
            catch (System.Exception ex)
            {
                failureReason = ex.Message;
                if (sideDb != null) sideDb.Dispose();
                return null;
            }
        }

        private static void EmbedNestedXrefsInSideDb(
            Database sideDb, Editor ed, HashSet<string> inFlightSourcePaths)
        {
            if (sideDb == null) return;

            var nestedDefIds = new List<ObjectId>();
            try
            {
                using (var tr = sideDb.TransactionManager.StartTransaction())
                {
                    var bt = tr.GetObject(sideDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt != null)
                    {
                        foreach (ObjectId btrId in bt)
                        {
                            var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                            if (btr == null) continue;
                            if (!IsDwgXrefDefinition(btr)) continue;
                            if (!btr.IsResolved) continue;
                            nestedDefIds.Add(btrId);
                        }
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\n    WARN: could not enumerate nested xrefs: {ex.Message}");
                return;
            }

            foreach (var nestedId in nestedDefIds)
            {
                var nested = EmbedSingleXref(sideDb, ed, nestedId, inFlightSourcePaths);
                if (!nested.Success)
                {
                    ed?.WriteMessage(
                        $"\n    WARN: nested embed of '{nested.XrefName}' failed " +
                        $"({nested.FailureCategory}): {nested.FailureReason}");
                }
            }
        }

        private static string GetUniqueLocalBlockName(Database hostDb, string xrefName)
        {
            string baseName = string.IsNullOrWhiteSpace(xrefName) ? "Xref" : xrefName.Trim();
            try { baseName = Path.GetFileNameWithoutExtension(baseName); } catch { }
            char[] illegal = { '<', '>', '/', '\\', '"', ':', ';', '?', '*', '|', '=', '`', ',' };
            foreach (var c in illegal) baseName = baseName.Replace(c, '_');
            if (string.IsNullOrWhiteSpace(baseName)) baseName = "Xref";

            string candidate = baseName + "_local";
            using (var tr = hostDb.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(hostDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (bt == null) { tr.Commit(); return candidate; }
                if (!bt.Has(candidate)) { tr.Commit(); return candidate; }
                int i = 2;
                while (bt.Has(candidate + "_" + i)) i++;
                tr.Commit();
                return candidate + "_" + i;
            }
        }

        // Uses DuplicateRecordCloning.Ignore so host's existing layer/linetype/style
        // records are kept (host-wins). Side-db records with no host-side name match
        // are cloned in fresh. This avoids needing a manual layer remap pass.
        private static int CloneModelspaceContents(
            Database hostDb, Database sideDb, ObjectId targetLocalBtrId,
            out string failureReason)
        {
            failureReason = string.Empty;
            try
            {
                ObjectId sideMsId = SymbolUtilityServices.GetBlockModelSpaceId(sideDb);
                if (sideMsId.IsNull)
                {
                    failureReason = "Side database has no modelspace BTR.";
                    return -1;
                }

                var sourceIds = new ObjectIdCollection();
                int skippedXrefRefs = 0;
                using (var sideTr = sideDb.TransactionManager.StartTransaction())
                {
                    var sideMs = sideTr.GetObject(sideMsId, OpenMode.ForRead) as BlockTableRecord;
                    if (sideMs == null)
                    {
                        sideTr.Commit();
                        failureReason = "Could not open side modelspace.";
                        return -1;
                    }
                    foreach (ObjectId id in sideMs)
                    {
                        if (!id.IsValid || id.IsErased) continue;

                        // Drop BlockReferences pointing to xref defs that weren't successfully
                        // embedded (e.g. failed nested embeds, cycles). Cloning them would
                        // propagate the xref def into the destination, leaving a def that
                        // can't be detached because the new local BTR holds an internal ref.
                        var ent = sideTr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent is BlockReference br)
                        {
                            var refBtr = sideTr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            if (refBtr != null && IsDwgXrefDefinition(refBtr))
                            {
                                skippedXrefRefs++;
                                continue;
                            }
                        }
                        sourceIds.Add(id);
                    }
                    sideTr.Commit();
                }

                if (skippedXrefRefs > 0)
                {
                    var ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage(
                        $"\n    Note: dropped {skippedXrefRefs} BlockReference(s) to unembedded xref def(s) (likely cycle/failure) to keep the host clean.");
                }

                if (sourceIds.Count == 0) return 0;

                var idMap = new IdMapping();
                hostDb.WblockCloneObjects(
                    sourceIds, targetLocalBtrId, idMap,
                    DuplicateRecordCloning.Ignore, false);
                return sourceIds.Count;
            }
            catch (System.Exception ex)
            {
                failureReason = ex.Message;
                return -1;
            }
        }

        private static int ReplaceModelspaceBlockReferences(
            Database hostDb, ObjectId oldXrefDefId, ObjectId newLocalBtrId,
            Transaction hostTr,
            out int paperspaceRefsRemaining,
            out string failureReason)
        {
            paperspaceRefsRemaining = 0;
            failureReason = string.Empty;

            try
            {
                ObjectId msId = SymbolUtilityServices.GetBlockModelSpaceId(hostDb);
                var defBtr = hostTr.GetObject(oldXrefDefId, OpenMode.ForRead) as BlockTableRecord;
                if (defBtr == null) { failureReason = "XREF definition is no longer valid."; return -1; }

                ObjectIdCollection refIds = defBtr.GetBlockReferenceIds(true, true);
                if (refIds == null || refIds.Count == 0) return 0;

                var modelspaceBtr = hostTr.GetObject(msId, OpenMode.ForWrite) as BlockTableRecord;
                if (modelspaceBtr == null) { failureReason = "Could not open modelspace."; return -1; }

                var layersToRelock = new HashSet<ObjectId>();
                int replaced = 0;

                foreach (ObjectId refId in refIds)
                {
                    if (!refId.IsValid || refId.IsErased) continue;
                    var oldBr = hostTr.GetObject(refId, OpenMode.ForRead) as BlockReference;
                    if (oldBr == null) continue;

                    if (oldBr.OwnerId != msId)
                    {
                        paperspaceRefsRemaining++;
                        continue;
                    }

                    var transform = oldBr.BlockTransform;
                    var position = oldBr.Position;
                    var layerId = oldBr.LayerId;
                    bool visible = oldBr.Visible;

                    if (layerId.IsValid && !layersToRelock.Contains(layerId))
                    {
                        var layer = hostTr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layer != null && layer.IsLocked)
                        {
                            var layerForWrite = hostTr.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                            if (layerForWrite != null)
                            {
                                layerForWrite.IsLocked = false;
                                layersToRelock.Add(layerId);
                            }
                        }
                    }

                    var newBr = new BlockReference(position, newLocalBtrId);
                    newBr.BlockTransform = transform;
                    newBr.LayerId = layerId;
                    newBr.Visible = visible;
                    ObjectId newBrId = modelspaceBtr.AppendEntity(newBr);
                    hostTr.AddNewlyCreatedDBObject(newBr, true);

                    PreserveDrawOrder(modelspaceBtr, refId, newBrId, hostTr);

                    var oldBrForWrite = hostTr.GetObject(refId, OpenMode.ForWrite) as BlockReference;
                    if (oldBrForWrite != null) oldBrForWrite.Erase();

                    replaced++;
                }

                foreach (var layerId in layersToRelock)
                {
                    var layer = hostTr.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                    if (layer != null) layer.IsLocked = true;
                }

                return replaced;
            }
            catch (System.Exception ex)
            {
                failureReason = ex.Message;
                return -1;
            }
        }

        private static void PreserveDrawOrder(
            BlockTableRecord ownerBtr, ObjectId oldBrId, ObjectId newBrId, Transaction tr)
        {
            try
            {
                if (!ownerBtr.DrawOrderTableId.IsValid) return;
                var dot = tr.GetObject(ownerBtr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                if (dot == null) return;
                dot.MoveAbove(new ObjectIdCollection(new[] { newBrId }), oldBrId);
            }
            catch { /* best-effort */ }
        }

        private static bool TryDetachIfFullyReplaced(
            Database hostDb, ObjectId xrefDefId, out string failureReason)
        {
            failureReason = string.Empty;
            try
            {
                int refCount;
                using (var tr = hostDb.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(xrefDefId, OpenMode.ForRead) as BlockTableRecord;
                    if (btr == null) { tr.Commit(); failureReason = "XREF def no longer valid."; return false; }
                    var ids = btr.GetBlockReferenceIds(true, true);
                    refCount = ids?.Count ?? 0;
                    tr.Commit();
                }

                if (refCount > 0)
                {
                    failureReason = $"{refCount} reference(s) still present.";
                    return false;
                }

                hostDb.DetachXref(xrefDefId);
                return true;
            }
            catch (System.Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
        }
    }
}
