using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("DETACHSPECIALXREFS", CommandFlags.Modal)]
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

                        // Also match if BOTH substrings "wl" and "sig" are present (case-insensitive)
                        if (!match)
                        {
                            bool nameWlSig = name.Contains("wl") && name.Contains("sig");
                            bool fileWlSig = !string.IsNullOrEmpty(fileNoExt) && fileNoExt.Contains("wl") && fileNoExt.Contains("sig");
                            if (nameWlSig || fileWlSig)
                                match = true;
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
                            // Also match if BOTH substrings "wl" and "sig" are present (case-insensitive)
                            if (!match)
                            {
                                bool keyWlSig = keyLower.Contains("wl") && keyLower.Contains("sig");
                                bool nameWlSig = nameLower.Contains("wl") && nameLower.Contains("sig");
                                bool nameNoExtWlSig = nameNoExtLower.Contains("wl") && nameNoExtLower.Contains("sig");
                                if (keyWlSig || nameWlSig || nameNoExtWlSig)
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
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId zeroId = ObjectId.Null;
                    LayerTableRecord zeroLtr = null;
                    if (lt.Has("0"))
                    {
                        zeroId = lt["0"];
                        zeroLtr = (LayerTableRecord)tr.GetObject(zeroId, OpenMode.ForWrite);
                        if (zeroLtr.IsOff) zeroLtr.IsOff = false;
                        if (zeroLtr.IsFrozen) zeroLtr.IsFrozen = false;
                    }

                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = null;
                        try { ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite); }
                        catch { continue; }
                        if (ltr == null) continue;

                        string lname = (ltr.Name ?? string.Empty).ToLowerInvariant();
                        bool matchChristian = lname.Contains("christian");
                        bool matchWlSig = lname.Contains("wl") && lname.Contains("sig");
                        bool matchWLstamp = lname.Contains("wlstamp");
                        bool matchRev = lname.Contains("rev");
                        bool matchDelta = lname.Contains("delta");
                        if ((matchChristian || matchWlSig || matchWLstamp || matchRev || matchDelta) && !ltr.IsFrozen)
                        {
                            try
                            {
                                if (db.Clayer == layerId)
                                {
                                    if (!zeroId.IsNull && zeroId != layerId)
                                    {
                                        db.Clayer = zeroId;
                                    }
                                    else
                                    {
                                        // If we cannot change, skip freezing current layer
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

                    tr.Commit();
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