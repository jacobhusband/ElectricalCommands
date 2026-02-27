using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("DETACHREMAININGXREFS", CommandFlags.Modal)]
        [CommandMethod("REMOVEREMAININGXREFS", CommandFlags.Modal)]
        public static void RemoveRemainingXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            int dwgDetached = 0;
            int imagesErased = 0;
            int imageDefsRemoved = 0;
            int pdfsErased = 0;
            int pdfDefsRemoved = 0;
            int blockRefsErased = 0;

            try
            {
                if (CleanupCommands.AbortRemainingXrefDetach)
                {
                    string reason = CleanupCommands.StrictTitleBlockBindFailed
                        ? "strict titleblock bind validation failed upstream"
                        : "upstream strict workflow requested an abort";
                    ed.WriteMessage($"\nDETACHREMAININGXREFS aborted: {reason}. No DWG/image/PDF references were removed.");
                    return;
                }

                var resolverLikelyIds = new HashSet<ObjectId>();
                if (CleanupCommands.StrictTitleBlockProtectionActive)
                {
                    foreach (var candidate in TitleBlockXrefResolver.GetLikelyTitleBlockCandidates(db))
                    {
                        if (!candidate.XrefBtrId.IsNull)
                        {
                            resolverLikelyIds.Add(candidate.XrefBtrId);
                        }
                    }
                }

                // --- Part 1: Detach ALL DWG XREFs ---
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var paperSpaceLayoutsByXref = BuildPaperSpaceLayoutMap(db, tr);
                    var protectedReasons = BuildProtectedTitleBlockCandidates(
                        db,
                        tr,
                        bt,
                        paperSpaceLayoutsByXref,
                        resolverLikelyIds);

                    var xrefIdsToDetach = new HashSet<ObjectId>();
                    if (protectedReasons.Count > 0)
                    {
                        ed.WriteMessage($"\nDETACHREMAININGXREFS: Protecting {protectedReasons.Count} titleblock-like XREF definition(s).");
                    }

                    // Collect all DWG XREF definitions
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null && btr.IsFromExternalReference)
                        {
                            if (protectedReasons.TryGetValue(btrId, out var reasons))
                            {
                                string reasonText = string.Join(", ", reasons.OrderBy(r => r, StringComparer.OrdinalIgnoreCase));
                                string shortPath = string.Empty;
                                try
                                {
                                    shortPath = Path.GetFileName(btr.PathName ?? string.Empty) ?? string.Empty;
                                }
                                catch { }
                                string pathSuffix = string.IsNullOrWhiteSpace(shortPath) ? string.Empty : $" ({shortPath})";
                                ed.WriteMessage($"\nSkipping protected titleblock XREF: {btr.Name}{pathSuffix} [{reasonText}]");
                                continue;
                            }
                            xrefIdsToDetach.Add(btrId);
                        }
                    }

                    if (xrefIdsToDetach.Count > 0)
                    {
                        // Erase all block references to these XREFs first
                        var refsToErase = new List<ObjectId>();
                        foreach (ObjectId spaceId in bt)
                        {
                            var spaceBtr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                            foreach (ObjectId entId in spaceBtr)
                            {
                                if (entId.ObjectClass.DxfName != "INSERT") continue;
                                var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                                if (br != null && xrefIdsToDetach.Contains(br.BlockTableRecord))
                                {
                                    refsToErase.Add(entId);
                                }
                            }
                        }

                        foreach (var entId in refsToErase)
                        {
                            try
                            {
                                var br = tr.GetObject(entId, OpenMode.ForWrite) as BlockReference;
                                if (br != null)
                                {
                                    var layer = (LayerTableRecord)tr.GetObject(br.LayerId, OpenMode.ForRead);
                                    bool relock = false;
                                    if (layer.IsLocked)
                                    {
                                        layer.UpgradeOpen();
                                        layer.IsLocked = false;
                                        relock = true;
                                    }

                                    br.Erase();
                                    blockRefsErased++;

                                    if (relock) layer.IsLocked = true;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nCould not erase block reference {entId}: {ex.Message}");
                            }
                        }

                        // Now detach the definitions
                        foreach (var xrefId in xrefIdsToDetach)
                        {
                            try
                            {
                                db.DetachXref(xrefId);
                                dwgDetached++;
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to detach DWG XREF {xrefId}: {ex.Message}");
                            }
                        }
                    }
                    tr.Commit();
                }

                // --- Part 2: Remove ALL Image References and Definitions ---
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // Step A: Erase all RasterImage entities from all spaces
                    var imagesToErase = new List<ObjectId>();
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId spaceId in bt)
                    {
                        var spaceBtr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                        foreach (ObjectId entId in spaceBtr)
                        {
                            if (entId.ObjectClass.DxfName == "IMAGE")
                            {
                                imagesToErase.Add(entId);
                            }
                        }
                    }

                    foreach (var imgId in imagesToErase)
                    {
                        try
                        {
                            var img = tr.GetObject(imgId, OpenMode.ForWrite) as RasterImage;
                            if (img != null && !img.IsErased)
                            {
                                var layer = (LayerTableRecord)tr.GetObject(img.LayerId, OpenMode.ForRead);
                                bool relock = false;
                                if (layer.IsLocked)
                                {
                                    layer.UpgradeOpen();
                                    layer.IsLocked = false;
                                    relock = true;
                                }
                                img.Erase();
                                imagesErased++;
                                if (relock) layer.IsLocked = true;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nCould not erase image {imgId}: {ex.Message}");
                        }
                    }

                    // Step B: Collect and remove all image definitions
                    var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (nod.Contains("ACAD_IMAGE_DICT"))
                    {
                        var imageDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);
                        var defsToRemove = new Dictionary<string, ObjectId>();

                        foreach (DBDictionaryEntry entry in imageDict)
                        {
                            defsToRemove.Add(entry.Key, entry.Value);
                        }

                        foreach (var kvp in defsToRemove)
                        {
                            try
                            {
                                imageDict.Remove(kvp.Key);
                                var def = tr.GetObject(kvp.Value, OpenMode.ForWrite);
                                def.Erase();
                                imageDefsRemoved++;
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to remove image def '{kvp.Key}': {ex.Message}");
                            }
                        }
                    }
                    tr.Commit();
                }

                // --- Part 3: Remove ALL PDF References and Definitions ---
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // Step A: Erase all PdfUnderlay entities from all spaces
                    var pdfsToErase = new List<ObjectId>();
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId spaceId in bt)
                    {
                        var spaceBtr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                        foreach (ObjectId entId in spaceBtr)
                        {
                            if (entId.ObjectClass.DxfName.Equals("PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                            {
                                pdfsToErase.Add(entId);
                            }
                        }
                    }

                    foreach (var pdfId in pdfsToErase)
                    {
                        try
                        {
                            var pdf = tr.GetObject(pdfId, OpenMode.ForWrite) as UnderlayReference;
                            if (pdf != null && !pdf.IsErased)
                            {
                                var layer = (LayerTableRecord)tr.GetObject(pdf.LayerId, OpenMode.ForRead);
                                bool relock = false;
                                if (layer.IsLocked)
                                {
                                    layer.UpgradeOpen();
                                    layer.IsLocked = false;
                                    relock = true;
                                }
                                pdf.Erase();
                                pdfsErased++;
                                if (relock) layer.IsLocked = true;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nCould not erase PDF underlay {pdfId}: {ex.Message}");
                        }
                    }

                    // Step B: Collect and remove all PDF definitions
                    var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (nod.Contains("ACAD_PDFDEFINITIONS"))
                    {
                        var pdfDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_PDFDEFINITIONS"), OpenMode.ForWrite);
                        var defsToRemove = new Dictionary<string, ObjectId>();

                        foreach (DBDictionaryEntry entry in pdfDict)
                        {
                            defsToRemove.Add(entry.Key, entry.Value);
                        }

                        foreach (var kvp in defsToRemove)
                        {
                            try
                            {
                                pdfDict.Remove(kvp.Key);
                                var def = tr.GetObject(kvp.Value, OpenMode.ForWrite);
                                def.Erase();
                                pdfDefsRemoved++;
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to remove PDF def '{kvp.Key}': {ex.Message}");
                            }
                        }
                    }
                    tr.Commit();
                }


                ed.WriteMessage($"\n--- XREF Removal Summary ---");
                ed.WriteMessage($"\nErased {blockRefsErased} DWG XREF block reference(s).");
                ed.WriteMessage($"\nDetached {dwgDetached} DWG XREF definition(s).");
                ed.WriteMessage($"\nErased {imagesErased} image reference(s).");
                ed.WriteMessage($"\nRemoved {imageDefsRemoved} image definition(s).");
                ed.WriteMessage($"\nErased {pdfsErased} PDF reference(s).");
                ed.WriteMessage($"\nRemoved {pdfDefsRemoved} PDF definition(s).");
                ed.WriteMessage($"\nCleanup complete.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during XREF removal: {ex.Message}");
            }
            finally
            {
                CleanupCommands.StrictTitleBlockBindFailed = false;
                CleanupCommands.AbortRemainingXrefDetach = false;
                CleanupCommands.ResetStrictTitleBlockProtection();
            }
        }

        private static Dictionary<ObjectId, HashSet<string>> BuildPaperSpaceLayoutMap(Database db, Transaction tr)
        {
            var result = new Dictionary<ObjectId, HashSet<string>>();

            try
            {
                var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (layoutDict == null) return result;

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead, false) as Layout;
                    if (layout == null || layout.ModelType) continue;

                    var psBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead, false) as BlockTableRecord;
                    if (psBtr == null) continue;

                    foreach (ObjectId entId in psBtr)
                    {
                        var br = tr.GetObject(entId, OpenMode.ForRead, false) as BlockReference;
                        if (br == null) continue;

                        var def = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead, false) as BlockTableRecord;
                        if (def == null || (!def.IsFromExternalReference && !def.IsFromOverlayReference)) continue;

                        if (!result.TryGetValue(br.BlockTableRecord, out var layoutNames))
                        {
                            layoutNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            result[br.BlockTableRecord] = layoutNames;
                        }
                        layoutNames.Add(layout.LayoutName ?? string.Empty);
                    }
                }
            }
            catch
            {
                // Best effort map for fallback matching.
            }

            return result;
        }

        private static Dictionary<ObjectId, HashSet<string>> BuildProtectedTitleBlockCandidates(
            Database db,
            Transaction tr,
            BlockTable bt,
            Dictionary<ObjectId, HashSet<string>> paperSpaceLayoutsByXref,
            HashSet<ObjectId> resolverLikelyIds)
        {
            var protectedReasons = new Dictionary<ObjectId, HashSet<string>>();
            bool strict = CleanupCommands.StrictTitleBlockProtectionActive;
            if (!strict) return protectedReasons;

            if (!CleanupCommands.ProtectedTitleBlockXrefId.IsNull)
            {
                AddProtectedReason(protectedReasons, CleanupCommands.ProtectedTitleBlockXrefId, "id");
            }

            if (resolverLikelyIds != null)
            {
                foreach (var id in resolverLikelyIds)
                {
                    if (!id.IsNull)
                    {
                        AddProtectedReason(protectedReasons, id, "resolver");
                    }
                }
            }

            foreach (ObjectId btrId in bt)
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead, false) as BlockTableRecord;
                if (btr == null || !btr.IsFromExternalReference) continue;

                bool inPaperSpace = paperSpaceLayoutsByXref.TryGetValue(btrId, out var layouts) && layouts.Count > 0;
                bool inProtectedLayout = inPaperSpace &&
                    !string.IsNullOrWhiteSpace(CleanupCommands.ProtectedTitleBlockLayoutName) &&
                    layouts.Contains(CleanupCommands.ProtectedTitleBlockLayoutName);

                bool fingerprintMatch = CleanupCommands.IsProtectedTitleBlockFingerprintMatch(db, btr.Name, btr.PathName);

                if (fingerprintMatch && (inPaperSpace || inProtectedLayout || string.IsNullOrWhiteSpace(CleanupCommands.ProtectedTitleBlockLayoutName)))
                {
                    AddProtectedReason(protectedReasons, btrId, "fingerprint");
                }

                if (protectedReasons.ContainsKey(btrId))
                {
                    if (inPaperSpace) AddProtectedReason(protectedReasons, btrId, "paper-space");
                    if (inProtectedLayout) AddProtectedReason(protectedReasons, btrId, "layout");
                }
            }

            return protectedReasons;
        }

        private static void AddProtectedReason(Dictionary<ObjectId, HashSet<string>> protectedReasons, ObjectId id, string reason)
        {
            if (id.IsNull || string.IsNullOrWhiteSpace(reason)) return;

            if (!protectedReasons.TryGetValue(id, out var reasons))
            {
                reasons = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                protectedReasons[id] = reasons;
            }

            reasons.Add(reason);
        }

        [CommandMethod("DETACHSIGX", CommandFlags.Modal)]
        [CommandMethod("QD", CommandFlags.Modal)]
        public static void DetachSpecialXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            // Case-insensitive tokens
            var imageTokens = new[] { "wl-sig", "christian" };
            var dwgTokens = new[] { "acieslogo", "cdstamp", "acies", "stamp", "logo" };

            int dwgDetached = 0;
            int imagesErased = 0;
            int imageDefsDetached = 0;
            int blockRefsErased = 0;
            int blockDefsErased = 0;
            int layersFrozen = 0;

            try
            {
                // Helper function for specific "wl-sig" pattern matching
                Func<string, bool> checkWlSig = s => s.Contains("wl-sig") || s.Contains("wl_sig") || s.Contains("wl sig");

                // 1) Detach DWG XREFs that match tokens or contain both "WL" and "sig"
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var toDetach = new List<ObjectId>();

                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null || !btr.IsFromExternalReference) continue;

                        string name = (btr.Name ?? string.Empty).ToLowerInvariant();
                        string path = (btr.PathName ?? string.Empty);
                        string fileNoExt = string.Empty;
                        try { fileNoExt = System.IO.Path.GetFileNameWithoutExtension(path)?.ToLowerInvariant() ?? string.Empty; } catch { }

                        bool match = false;
                        foreach (var t in dwgTokens)
                        {
                            var token = t.ToLowerInvariant();
                            if (name.Contains(token) || (!string.IsNullOrEmpty(fileNoExt) && fileNoExt.Contains(token)))
                            {
                                match = true;
                                break;
                            }
                        }

                        // Also match if a specific "wl-sig" pattern is present
                        if (!match)
                        {
                            if (checkWlSig(name) || (!string.IsNullOrEmpty(fileNoExt) && checkWlSig(fileNoExt)))
                            {
                                match = true;
                            }
                        }

                        if (match)
                        {
                            toDetach.Add(btrId);
                        }
                    }

                    // Erase all block references to the XREFs before detaching
                    var refsToErase = new List<ObjectId>();
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr.IsLayout) continue;

                        foreach (ObjectId entId in btr)
                        {
                            if (entId.ObjectClass.DxfName != "INSERT") continue;
                            var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;

                            if (toDetach.Contains(br.BlockTableRecord))
                            {
                                refsToErase.Add(entId);
                            }
                        }
                    }

                    foreach (var entId in refsToErase)
                    {
                        try
                        {
                            var br = tr.GetObject(entId, OpenMode.ForWrite, false) as BlockReference;
                            if (br != null)
                            {
                                br.Erase();
                                blockRefsErased++;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nFailed erasing block reference {entId}: {ex.Message}");
                        }
                    }

                    foreach (var xrefId in toDetach)
                    {
                        try { db.DetachXref(xrefId); dwgDetached++; }
                        catch (System.Exception ex) { ed.WriteMessage($"\nFailed to detach DWG XREF {xrefId}: {ex.Message}"); }
                    }

                    tr.Commit();
                }

                // 2) Detach image references/defs that match tokens
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!nod.Contains("ACAD_IMAGE_DICT"))
                    {
                        tr.Commit();
                    }
                    else
                    {
                        var imageDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);

                        // Gather candidate defs by key or source file name
                        var candidates = new List<(string Key, ObjectId DefId)>();
                        foreach (DBDictionaryEntry entry in imageDict)
                        {
                            string key = entry.Key ?? string.Empty;
                            var def = tr.GetObject(entry.Value, OpenMode.ForRead) as RasterImageDef;
                            string src = def?.SourceFileName ?? string.Empty;
                            string keyLower = key.ToLowerInvariant();
                            string nameLower = string.Empty;
                            string nameNoExtLower = string.Empty;
                            try { nameLower = System.IO.Path.GetFileName(src)?.ToLowerInvariant() ?? string.Empty; } catch { }
                            try { nameNoExtLower = System.IO.Path.GetFileNameWithoutExtension(src)?.ToLowerInvariant() ?? string.Empty; } catch { }

                            bool match = false;
                            foreach (var t in imageTokens)
                            {
                                var token = t.ToLowerInvariant();
                                if (keyLower.Contains(token) || nameLower.Contains(token) || nameNoExtLower.Contains(token))
                                {
                                    match = true;
                                    break;
                                }
                            }
                            // Also match if a specific "wl-sig" pattern is present
                            if (!match)
                            {
                                if (checkWlSig(keyLower) || checkWlSig(nameLower) || checkWlSig(nameNoExtLower))
                                {
                                    match = true;
                                }
                            }
                            if (match)
                            {
                                candidates.Add((key, entry.Value));
                            }
                        }

                        if (candidates.Count > 0)
                        {
                            // Build an index of raster images by def id
                            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var imagesByDef = new Dictionary<ObjectId, List<ObjectId>>();
                            foreach (var c in candidates)
                                imagesByDef[c.DefId] = new List<ObjectId>();

                            foreach (ObjectId btrId in bt)
                            {
                                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                foreach (ObjectId entId in btr)
                                {
                                    if (entId.ObjectClass.DxfName != "IMAGE") continue;
                                    var img = tr.GetObject(entId, OpenMode.ForRead) as RasterImage;
                                    if (img == null || img.ImageDefId.IsNull) continue;
                                    if (imagesByDef.ContainsKey(img.ImageDefId))
                                    {
                                        imagesByDef[img.ImageDefId].Add(entId);
                                    }
                                }
                            }

                            // Erase image references and then remove defs
                            foreach (var c in candidates)
                            {
                                // Erase all image entities referencing this def
                                if (imagesByDef.TryGetValue(c.DefId, out var entIds))
                                {
                                    foreach (var entId in entIds)
                                    {
                                        try
                                        {
                                            var img = tr.GetObject(entId, OpenMode.ForWrite, false) as RasterImage;
                                            if (img != null)
                                            {
                                                // Unlock layer if needed
                                                var layer = (LayerTableRecord)tr.GetObject(img.LayerId, OpenMode.ForRead);
                                                bool relock = false;
                                                if (layer.IsLocked)
                                                {
                                                    layer.UpgradeOpen();
                                                    layer.IsLocked = false;
                                                    relock = true;
                                                }
                                                img.Erase();
                                                imagesErased++;
                                                if (relock) layer.IsLocked = true;
                                            }
                                        }
                                        catch (System.Exception ex)
                                        {
                                            ed.WriteMessage($"\nFailed erasing image ref {entId}: {ex.Message}");
                                        }
                                    }
                                }

                                // Remove dictionary key and erase def
                                try { imageDict.Remove(c.Key); } catch { }
                                try
                                {
                                    var def = tr.GetObject(c.DefId, OpenMode.ForWrite, false);
                                    def?.Erase();
                                    imageDefsDetached++;
                                }
                                catch (System.Exception ex)
                                {
                                    ed.WriteMessage($"\nFailed to detach image def '{c.Key}': {ex.Message}");
                                }
                            }
                        }

                        tr.Commit();
                    }
                }

                // 3) Delete block references and block definitions matching patterns
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // Identify matching block definitions (non-layout, non-xref)
                    var matchedDefs = new HashSet<ObjectId>();
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null) continue;
                        if (btr.IsLayout || btr.IsDependent || btr.IsFromExternalReference) continue;

                        // Case-insensitive match for block name containing "STAMP"
                        string bname = (btr.Name ?? string.Empty).ToLowerInvariant();
                        if (bname.Contains("stamp"))
                        {
                            matchedDefs.Add(btrId);
                        }
                    }

                    if (matchedDefs.Count > 0)
                    {
                        // Collect all block references to these definitions
                        var refsToErase = new List<ObjectId>();
                        foreach (ObjectId spaceId in bt)
                        {
                            var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                            foreach (ObjectId entId in btr)
                            {
                                if (entId.ObjectClass.DxfName != "INSERT") continue;
                                var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                                if (br == null) continue;

                                ObjectId blockId = br.BlockTableRecord;
                                if (br.IsDynamicBlock)
                                {
                                    try { blockId = br.DynamicBlockTableRecord; }
                                    catch { }
                                }

                                if (matchedDefs.Contains(blockId))
                                {
                                    refsToErase.Add(entId);
                                }
                            }
                        }

                        // Erase all matching references
                        foreach (var entId in refsToErase)
                        {
                            try
                            {
                                var br = tr.GetObject(entId, OpenMode.ForWrite, false) as BlockReference;
                                if (br != null)
                                {
                                    br.Erase();
                                    blockRefsErased++;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed erasing block reference {entId}: {ex.Message}");
                            }
                        }

                        // Erase the block definitions now that refs are gone
                        foreach (var defId in matchedDefs)
                        {
                            try
                            {
                                var btr = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForWrite, false);
                                if (btr != null)
                                {
                                    btr.Erase();
                                    blockDefsErased++;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed erasing block definition {defId}: {ex.Message}");
                            }
                        }
                    }

                    tr.Commit();
                }

                // 4) Freeze layers matching requested patterns
                if (!_skipLayerFreezing)
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                        ObjectId originalClayer = db.Clayer;
                        ObjectId fallbackClayer = ObjectId.Null;

                        foreach (ObjectId layerId in lt)
                        {
                            LayerTableRecord ltr = null;
                            try { ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite); }
                            catch { continue; }
                            if (ltr == null) continue;

                            string lname = (ltr.Name ?? string.Empty).ToLowerInvariant();
                            bool matchCloud = lname.Contains("cloud");
                            bool matchDelta = lname.Contains("delta");
                            if ((matchCloud || matchDelta) && !ltr.IsFrozen)
                            {
                                try
                                {
                                    if (db.Clayer == layerId)
                                    {
                                        if (fallbackClayer.IsNull)
                                        {
                                            foreach (ObjectId candidateId in lt)
                                            {
                                                if (candidateId == layerId) continue;
                                                var candidate = tr.GetObject(candidateId, OpenMode.ForRead, false) as LayerTableRecord;
                                                if (candidate == null || candidate.IsErased || candidate.IsFrozen || candidate.IsOff) continue;
                                                fallbackClayer = candidateId;
                                                break;
                                            }
                                        }
                                        if (!fallbackClayer.IsNull)
                                        {
                                            db.Clayer = fallbackClayer;
                                        }
                                        else
                                        {
                                            ed.WriteMessage($"\nSkipping freeze for current layer '{ltr.Name}' (no safe fallback current layer available).");
                                            continue;
                                        }
                                    }
                                    ltr.IsFrozen = true;
                                    layersFrozen++;
                                }
                                catch (System.Exception ex)
                                {
                                    ed.WriteMessage($"\nFailed to freeze layer '{ltr.Name}': {ex.Message}");
                                }
                            }
                        }

                        try
                        {
                            if (!originalClayer.IsNull &&
                                originalClayer.IsValid &&
                                !originalClayer.IsErased)
                            {
                                var originalLtr = tr.GetObject(originalClayer, OpenMode.ForRead, false) as LayerTableRecord;
                                if (originalLtr != null && !originalLtr.IsErased && !originalLtr.IsFrozen)
                                {
                                    db.Clayer = originalClayer;
                                }
                            }
                        }
                        catch { }

                        tr.Commit();
                    }
                }
                else
                {
                    ed.WriteMessage("\nSkipping layer freezing as requested by the current workflow.");
                }

                ed.WriteMessage($"\nDetached {dwgDetached} DWG XREF(s). Erased {imagesErased} image ref(s) and detached {imageDefsDetached} image def(s). Erased {blockRefsErased} block ref(s) and deleted {blockDefsErased} block def(s). Froze {layersFrozen} layer(s).");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nDetach operation failed: {ex.Message}");
            }
        }
    }
}
