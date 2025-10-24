using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCADCleanupTool
{
    public class PreExplosionCleaner
    {
        /// <summary>
        /// This command method orchestrates the entire cleanup process as requested.
        /// 1. Detaches specific XREFs and images.
        /// 2. Erases specific block references and their definitions.
        /// 3. Explodes all remaining block references.
        /// 4. Erases any remaining objects on signature, logo, or stamp-related layers.
        /// </summary>
        [CommandMethod("CLEAN_BEFORE_EXPLODE", CommandFlags.Modal)]
        public static void CleanBeforeExplode()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage("\n--- Starting Pre-Explosion Cleanup ---");

            try
            {
                // Preparation: Ensure all layers are visible and editable.
                EnsureAllLayersVisibleAndUnlocked(db, ed);

                // Step 1 & 2: Perform targeted cleanup of specific XREFs and blocks.
                PerformTargetedCleanup(db, ed);

                // Step 3: Explode all remaining block references to expose nested entities.
                ExplodeAllBlockReferences(db, ed);

                // Step 4: Perform a robust, layer-based cleanup of any remaining stray objects.
                // --- MODIFIED: Added "WLSIG" and "Christian" to the layer keywords ---
                string[] layerKeywords = { "SIG", "LOGO", "STAMP", "ACIES", "CDSTAMP", "WLSIG", "Christian" };
                EraseObjectsOnMatchingLayers(db, ed, layerKeywords);

                ed.WriteMessage("\n--- Pre-Explosion Cleanup Completed Successfully ---");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during the cleanup process: {ex.Message}");
            }
        }

        /// <summary>
        /// Detaches specific DWG/Image XREFs and erases specific block definitions and their references.
        /// </summary>
        private static void PerformTargetedCleanup(Database db, Editor ed)
        {
            ed.WriteMessage("\nPerforming targeted cleanup of XREFs and blocks...");

            // --- Define the specific names to target ---
            var dwgXrefsToDetach = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "CDstamp-CA CAD XREF" };
            var imageNamesToDetach = new[] { "CD_sig_BLUE", "WL_Sig_Blue_Small" };

            // --- MODIFIED: Added "x-Christian's signature$0$PFV" to the blocks to erase ---
            var blocksToErase = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "STAMP_WL",
                "ACIESLOGO",
                "x-Christian's signature$0$PFV"
            };

            int dwgDetached = 0, imagesDetached = 0, blocksErased = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // --- Part 1: Handle DWG XREFs and regular Blocks ---
                var blockDefIdsToClean = new List<ObjectId>();
                foreach (ObjectId btrId in bt)
                {
                    var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                    if (btr == null) continue;

                    // Match DWG XREFs to detach
                    if (btr.IsFromExternalReference && dwgXrefsToDetach.Contains(btr.Name))
                    {
                        blockDefIdsToClean.Add(btrId);
                    }
                    // Match regular blocks to erase
                    else if (!btr.IsFromExternalReference && blocksToErase.Contains(btr.Name))
                    {
                        blockDefIdsToClean.Add(btrId);
                    }
                }

                // Erase all references before attempting to detach/erase definitions
                if (blockDefIdsToClean.Count > 0)
                {
                    EraseBlockReferences(tr, bt, blockDefIdsToClean);

                    foreach (var defId in blockDefIdsToClean)
                    {
                        try
                        {
                            var btr = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForRead);
                            if (btr.IsFromExternalReference)
                            {
                                db.DetachXref(defId);
                                dwgDetached++;
                            }
                            else
                            {
                                btr.UpgradeOpen();
                                btr.Erase();
                                blocksErased++;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nFailed to process block/XREF definition {defId}: {ex.Message}");
                        }
                    }
                }

                // --- Part 2: Handle Image XREFs ---
                var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                if (nod.Contains("ACAD_IMAGE_DICT"))
                {
                    var imageDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);
                    var imageDefsToDetach = new Dictionary<string, ObjectId>();

                    foreach (DBDictionaryEntry entry in imageDict)
                    {
                        var def = tr.GetObject(entry.Value, OpenMode.ForRead) as RasterImageDef;
                        if (def == null) continue;

                        string sourceFile = System.IO.Path.GetFileNameWithoutExtension(def.SourceFileName);
                        if (imageNamesToDetach.Any(name => sourceFile.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            imageDefsToDetach.Add(entry.Key, entry.Value);
                        }
                    }

                    foreach (var kvp in imageDefsToDetach)
                    {
                        EraseImageReferences(tr, bt, kvp.Value);
                        try
                        {
                            imageDict.Remove(kvp.Key);
                            var def = tr.GetObject(kvp.Value, OpenMode.ForWrite);
                            def.Erase();
                            imagesDetached++;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nFailed to detach image definition '{kvp.Key}': {ex.Message}");
                        }
                    }
                }

                tr.Commit();
            }
            ed.WriteMessage($"\nCleanup Summary: Detached {dwgDetached} DWG XREF(s), {imagesDetached} Image(s), and erased {blocksErased} Block Definition(s).");
        }

        #region Helper Methods

        /// <summary>
        /// Erases all BlockReference instances that point to a list of BlockTableRecord definitions.
        /// </summary>
        private static void EraseBlockReferences(Transaction tr, BlockTable bt, List<ObjectId> blockDefIds)
        {
            var defIdSet = new HashSet<ObjectId>(blockDefIds);
            var modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId entId in modelSpace)
            {
                if (entId.ObjectClass.DxfName == "INSERT")
                {
                    var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                    if (br != null && defIdSet.Contains(br.BlockTableRecord))
                    {
                        br.UpgradeOpen();
                        br.Erase();
                    }
                }
            }
        }

        /// <summary>
        /// Erases all RasterImage instances that point to a specific RasterImageDef.
        /// </summary>
        private static void EraseImageReferences(Transaction tr, BlockTable bt, ObjectId imageDefId)
        {
            var modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId entId in modelSpace)
            {
                if (entId.ObjectClass.DxfName == "IMAGE")
                {
                    var img = tr.GetObject(entId, OpenMode.ForRead) as RasterImage;
                    if (img != null && img.ImageDefId == imageDefId)
                    {
                        img.UpgradeOpen();
                        img.Erase();
                    }
                }
            }
        }

        /// <summary>
        /// Iteratively explodes all block references in the Model Space until none are left.
        /// </summary>
        private static void ExplodeAllBlockReferences(Database db, Editor ed)
        {
            ed.WriteMessage("\nExploding all remaining block references...");
            int totalExploded = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
                int explodedInPass;
                do
                {
                    explodedInPass = 0;
                    var blockRefsToExplode = new List<ObjectId>();
                    foreach (ObjectId id in modelSpace)
                    {
                        if (id.ObjectClass.DxfName == "INSERT")
                        {
                            blockRefsToExplode.Add(id);
                        }
                    }

                    if (blockRefsToExplode.Count > 0)
                    {
                        foreach (var id in blockRefsToExplode)
                        {
                            try
                            {
                                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                                var explodedObjects = new DBObjectCollection();
                                br.Explode(explodedObjects);

                                foreach (DBObject obj in explodedObjects)
                                {
                                    var ent = obj as Entity;
                                    if (ent != null)
                                    {
                                        modelSpace.AppendEntity(ent);
                                        tr.AddNewlyCreatedDBObject(ent, true);
                                    }
                                }
                                br.UpgradeOpen();
                                br.Erase();
                                explodedInPass++;
                            }
                            catch { /* Ignore non-explodable blocks */ }
                        }
                        totalExploded += explodedInPass;
                    }
                } while (explodedInPass > 0);
                tr.Commit();
            }
            ed.WriteMessage($"\nExploded a total of {totalExploded} block references.");
        }

        /// <summary>
        /// Erases all entities in Model Space that reside on a layer whose name contains any of the specified keywords.
        /// </summary>
        private static void EraseObjectsOnMatchingLayers(Database db, Editor ed, string[] layerKeywords)
        {
            ed.WriteMessage("\nPerforming robust cleanup of objects on specified layers...");
            int objectsErased = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                var layersToClean = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Find all layers that match the keywords
                foreach (ObjectId layerId in lt)
                {
                    var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    if (layerKeywords.Any(kw => ltr.Name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        layersToClean.Add(ltr.Name);
                    }
                }

                if (layersToClean.Count > 0)
                {
                    foreach (ObjectId id in modelSpace)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent != null && layersToClean.Contains(ent.Layer))
                        {
                            ent.UpgradeOpen();
                            ent.Erase();
                            objectsErased++;
                        }
                    }
                }
                tr.Commit();
            }
            ed.WriteMessage($"\nErased {objectsErased} objects found on matching layers.");
        }

        /// <summary>
        /// Unlocks, thaws, and turns on all layers in the drawing.
        /// </summary>
        private static void EnsureAllLayersVisibleAndUnlocked(Database db, Editor ed)
        {
            ed.WriteMessage("\nEnsuring all layers are visible and unlocked...");
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId layerId in lt)
                {
                    if (layerId.IsErased) continue;
                    var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    if (ltr.IsLocked) ltr.IsLocked = false;
                    if (ltr.IsFrozen) ltr.IsFrozen = false;
                    if (ltr.IsOff) ltr.IsOff = false;
                }
                tr.Commit();
            }
        }

        #endregion
    }
}