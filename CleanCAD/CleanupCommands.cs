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
    public class CleanupCommands
    {
        // Static variables to pass state between commands
        private static HashSet<ObjectId> _blockIdsBeforeBind = new HashSet<ObjectId>();
        private static HashSet<ObjectId> _originalXrefIds = new HashSet<ObjectId>();
        private static HashSet<ObjectId> _imageDefsToPurge = new HashSet<ObjectId>(); // Our definitive "kill list"

        // STAGE 1: Get snapshots, bind, and queue the main cleanup command.
        [CommandMethod("FINALIZE")]
        public static void FinalizeDrawingCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            _blockIdsBeforeBind.Clear();
            _originalXrefIds.Clear();
            _imageDefsToPurge.Clear();
            ed.WriteMessage("\n--- Stage 1: Analyzing and Binding... ---");

            try
            {
                int bindCount = 0;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    foreach (ObjectId btrId in bt)
                    {
                        _blockIdsBeforeBind.Add(btrId);
                        var btr = trans.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null && btr.IsFromExternalReference && btr.IsResolved &&
                            !string.IsNullOrEmpty(btr.PathName) && btr.PathName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
                        {
                            _originalXrefIds.Add(btrId);
                        }
                    }
                    ed.WriteMessage($"\nFound {_blockIdsBeforeBind.Count} total blocks and {_originalXrefIds.Count} DWG XREFs to bind.");

                    if (_originalXrefIds.Count > 0)
                    {
                        var idsToBind = new ObjectIdCollection(_originalXrefIds.ToArray());
                        ed.WriteMessage($"\nBinding {idsToBind.Count} DWG reference(s)...");
                        db.BindXrefs(idsToBind, true);
                        bindCount = idsToBind.Count;
                    }
                    trans.Commit();
                }

                if (bindCount > 0)
                {
                    ed.WriteMessage("\nBind complete. Queueing cleanup process...");
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
            }
        }

        // STAGE 2: Erase images, build the "kill list", and detach ghost XREF blocks.
        [CommandMethod("-FINALIZE-CLEANUP", CommandFlags.Modal)]
        public static void FinalizeCleanupCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n--- Stage 2: Erasing images and detaching ghost XREF blocks... ---");

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
                    int imagesEmbedded = 0;
                    foreach (ObjectId newBtrId in newBlockIds)
                    {
                        BlockTableRecord newBtr = trans.GetObject(newBtrId, OpenMode.ForRead) as BlockTableRecord;
                        foreach (ObjectId entId in newBtr)
                        {
                            if (entId.ObjectClass.DxfName == "IMAGE")
                            {
                                RasterImage image = trans.GetObject(entId, OpenMode.ForWrite) as RasterImage;
                                if (image == null)
                                    continue;

                                if (image.ImageDefId.IsNull)
                                {
                                    ed.WriteMessage("\n  -> WARNING: IMAGE entity has no ImageDef; skipping.");
                                    continue;
                                }

                                // Open the image definition to get its source path
                                RasterImageDef imageDef = trans.GetObject(image.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                                if (imageDef == null)
                                {
                                    ed.WriteMessage("\n  -> WARNING: Could not open RasterImageDef; skipping.");
                                    continue;
                                }

                                string imagePath = imageDef.SourceFileName;
                                if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                                {
                                    ed.WriteMessage($"\n  -> WARNING: Source image not found, cannot embed: {imagePath}");
                                    continue;
                                }

                                Point3d position = image.Position;
                                double rotation = image.Rotation; // Already in radians
                                double scale = image.Scale.X;     // Assume uniform scale

                                ed.WriteMessage($"\n  -> Found linked image: {imagePath}");

                                bool success = EmbedImageViaLisp(imagePath, position, scale, rotation);

                                if (success)
                                {
                                    ed.WriteMessage("     ...Success: Embedded a copy via LISP interop.");
                                    imagesEmbedded++;
                                    _imageDefsToPurge.Add(image.ImageDefId); // Queue original def for purge

                                    image.Erase();
                                    imagesErased++;
                                }
                                else
                                {
                                    ed.WriteMessage("     ...FAILURE: Could not embed. The linked image will be left in place.");
                                }
                            }
                        }
                    }
                    ed.WriteMessage($"\nEmbedded {imagesEmbedded} image(s). Erased {imagesErased} RasterImage entit(ies). Added {_imageDefsToPurge.Count} definitions to kill list.");

                    int ghostsDetached = 0;
                    foreach (ObjectId originalXrefId in _originalXrefIds)
                    {
                        var obj = trans.GetObject(originalXrefId, OpenMode.ForRead, false, true);
                        if (obj != null && !obj.IsErased)
                        {
                            db.DetachXref(originalXrefId);
                            ghostsDetached++;
                        }
                    }
                    ed.WriteMessage($"\nManually detached {ghostsDetached} old XREF block definition(s).");

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
                _blockIdsBeforeBind.Clear();
                _originalXrefIds.Clear();
            }
        }

        // STAGE 3: Surgically remove the image definitions from our kill list.
        [CommandMethod("-FINALIZE-PURGEDEFS", CommandFlags.Modal)]
        public static void FinalizePurgeDefsCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n--- Stage 3: Surgically removing orphaned image definitions... ---");
            if (_imageDefsToPurge.Count == 0)
            {
                ed.WriteMessage("\nNo orphaned image definitions were identified. Finalization complete.");
                return;
            }

            try
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    int defsPurged = 0;
                    DBDictionary namedObjectsDict = trans.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (namedObjectsDict.Contains("ACAD_IMAGE_DICT"))
                    {
                        DBDictionary imageDict = trans.GetObject(namedObjectsDict.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite) as DBDictionary;
                        var entriesToRemove = new Dictionary<string, ObjectId>();

                        // Find the dictionary keys for the defs on our kill list
                        foreach (DBDictionaryEntry entry in imageDict)
                        {
                            if (_imageDefsToPurge.Contains(entry.Value))
                            {
                                entriesToRemove.Add(entry.Key, entry.Value);
                            }
                        }

                        ed.WriteMessage($"\nFound {entriesToRemove.Count} dictionary entr(ies) to remove.");
                        foreach (var item in entriesToRemove)
                        {
                            ed.WriteMessage($"\n  - Detaching and Erasing: {item.Key}");
                            imageDict.Remove(item.Key);
                            DBObject imageDef = trans.GetObject(item.Value, OpenMode.ForWrite);
                            imageDef.Erase();
                            defsPurged++;
                        }
                    }
                    ed.WriteMessage($"\nSuccessfully purged {defsPurged} image definition(s).");
                    trans.Commit();
                }
                ed.WriteMessage("\n-------------------------------------------------");
                ed.WriteMessage("\nFinalization Complete. Please save the drawing.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during final purge: {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                _imageDefsToPurge.Clear();
            }
        }

        // LISP Interop: Call ARX-exposed function to embed an image
        private static bool EmbedImageViaLisp(string imagePath, Point3d position, double scale, double rotationInRadians)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return false;

            try
            {
                ResultBuffer args = new ResultBuffer();
                // LISP function name
                args.Add(new TypedValue((int)LispDataType.Text, "C:EmbedImage-ARX"));
                // Argument 1: Image Path
                args.Add(new TypedValue((int)LispDataType.Text, imagePath));
                // Argument 2: Insertion Point
                args.Add(new TypedValue((int)LispDataType.Point3d, position));
                // Argument 3: Scale
                args.Add(new TypedValue((int)LispDataType.Double, scale));
                // Argument 4: Rotation in Radians
                args.Add(new TypedValue((int)LispDataType.Double, rotationInRadians));

                Application.Invoke(args);

                return true; // Assume success if no exception is thrown
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\nLISP interop failed: {ex.Message}");
                return false;
            }
        }
    }
}


