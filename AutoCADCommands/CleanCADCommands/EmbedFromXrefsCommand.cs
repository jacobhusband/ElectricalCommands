using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        /// <summary>
        /// Utility: explode all block references in Model Space (unchanged).
        /// </summary>
        public static void ExplodeAllBlockReferences()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    int explodedCount;
                    do
                    {
                        explodedCount = 0;
                        var blockReferences = new List<BlockReference>();

                        foreach (ObjectId id in modelSpace)
                        {
                            if (id.ObjectClass.DxfName.Equals("INSERT", StringComparison.OrdinalIgnoreCase))
                            {
                                var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                                if (br != null)
                                    blockReferences.Add(br);
                            }
                        }

                        if (blockReferences.Count > 0)
                        {
                            ed.WriteMessage($"\nFound {blockReferences.Count} block references to explode...");
                            foreach (var br in blockReferences)
                            {
                                var explodedObjects = new DBObjectCollection();
                                br.UpgradeOpen();
                                br.Explode(explodedObjects);

                                foreach (DBObject obj in explodedObjects)
                                {
                                    if (obj is Entity ent)
                                    {
                                        modelSpace.AppendEntity(ent);
                                        tr.AddNewlyCreatedDBObject(ent, true);
                                    }
                                }
                                br.Erase();
                                explodedCount++;
                            }
                            ed.WriteMessage($"\nExploded {explodedCount} block references.");
                        }

                    } while (explodedCount > 0);

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError while exploding block references: {ex.Message}");
            }
        }

        /// <summary>
        /// EMBEDIMAGES:
        /// - Collects all raster images (including from XREFs) across layouts.
        /// - Pre-rotates each image to match its orientation.
        /// - Uses PowerPoint as an OLE source and pastes into the correct layout.
        /// - Transforms the pasted OLE to align with the original image's oriented rectangle.
        /// - Records XREFs for detaching and image defs for purging.
        /// This implementation mirrors your previously working rotated-raster behavior.
        /// </summary>
        [CommandMethod("EMBEDIMAGES", CommandFlags.Modal)]
        public static void EmbedImages()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage("\n--- Starting EMBEDIMAGES ---");

            _pendingXref.Clear();
            _lastPastedOle = ObjectId.Null;
            _isEmbeddingProcessActive = false;
            _finalPastedOleForZoom = ObjectId.Null;
            _xrefsToDetach.Clear(); // Ensure clean slate

            try
            {
                DeleteOldEmbedTemps(7);
            }
            catch { /* Ignore temp clean errors */ }

            ObjectId originalClayer = db.Clayer;
            Matrix3d originalUcs = ed.CurrentUserCoordinateSystem;
            string originalLayout = LayoutManager.Current.CurrentLayout;

            try
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;

                // 1. THAW LAYER 0
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has("0"))
                    {
                        var zeroId = lt["0"];
                        var zeroLtr = (LayerTableRecord)tr.GetObject(zeroId, OpenMode.ForWrite);
                        if (zeroLtr.IsFrozen) zeroLtr.IsFrozen = false;
                        if (_savedClayer.IsNull) _savedClayer = originalClayer;
                        db.Clayer = zeroId;
                    }
                    tr.Commit();
                }

                RemoveInvalidRasterImages(db, ed);

                // 2. EXPLODE BLOCKS FIRST
                // If the JPG is inside an XREF or Block, the scanner won't find it unless we explode it first.
                ed.WriteMessage("\nPre-exploding blocks to expose nested images...");
                ExplodeAllBlockReferences();

                // 3. COLLECT IMAGES
                try
                {
                    ed.WriteMessage("\nScanning all layouts for raster images...");
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                        // Scan Model
                        CollectAndPreflightImagesInLayoutForXrefs(doc, bt[BlockTableRecord.ModelSpace]);

                        // Scan Layouts
                        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            if (entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;
                            var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                            CollectAndPreflightImagesInLayoutForXrefs(doc, layout.BlockTableRecordId);
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[!] Critical error during collection: {ex.Message}");
                    return;
                }

                if (_pendingXref.Count == 0)
                {
                    ed.WriteMessage("\nNo valid raster images found to embed.");
                    FinishEmbeddingRun(doc, ed, db);
                    return;
                }

                ed.WriteMessage($"\nSuccessfully queued {_pendingXref.Count} raster image(s).");

                if (!EnsurePowerPoint(ed)) return;

                if (WindowOrchestrator.TryGetPowerPointHwnd(out var pptHwnd))
                    WindowOrchestrator.EnsureSeparationOrSafeOverlap(ed, pptHwnd, preferDifferentMonitor: true);

                AttachHandlersForXrefs(db, doc);
                _isEmbeddingProcessActive = true;
                ProcessNextPasteForXrefs(doc, ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[!] Error: {ex.Message}");
                FinishEmbeddingRun(doc, ed, db); // Ensure cleanup happens even on crash
            }
        }

        /// <summary>
        /// Collect and preflight RasterImage entities for the XREF pipeline.
        /// Uses XREF-specific ImagePlacementXref and queues into _pendingXref.
        /// </summary>
        private static void CollectAndPreflightImagesInLayoutForXrefs(Document doc, ObjectId spaceId)
        {
            var db = doc.Database;
            var ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var space = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);

                foreach (ObjectId id in space)
                {
                    // Look for raster images
                    if (!id.ObjectClass.DxfName.Equals("IMAGE", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var img = tr.GetObject(id, OpenMode.ForRead) as RasterImage;
                    if (img == null || img.ImageDefId.IsNull) continue;

                    var def = tr.GetObject(img.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                    if (def == null) continue;

                    // Mark Owner XREF for detach
                    if (tr.GetObject(img.OwnerId, OpenMode.ForRead) is BlockTableRecord ownerBtr &&
                        ownerBtr.IsFromExternalReference)
                    {
                        _xrefsToDetach.Add(ownerBtr.ObjectId);
                    }

                    // Resolve Path
                    string resolved = ResolveImagePath(db, def.SourceFileName);
                    if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                    {
                        // Try to find it in the same folder as the DWG if path resolution failed
                        string localAttempt = Path.Combine(Path.GetDirectoryName(doc.Name), Path.GetFileName(def.SourceFileName));
                        if (File.Exists(localAttempt))
                        {
                            resolved = localAttempt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n[Skipping] Missing image file: {def.SourceFileName}");
                            continue;
                        }
                    }

                    string safePath = resolved;
                    var cs = img.Orientation;
                    double rotationDeg = 0.0;

                    // --- CMYK / PREFLIGHT FIX ---
                    try
                    {
                        // Calculate AutoCAD rotation
                        rotationDeg = Math.Atan2(cs.Xaxis.Y, cs.Xaxis.X) * (180.0 / Math.PI);

                        // Attempt Preflight (Checks dimensions, rotation, etc.)
                        // If this is a CMYK JPG, System.Drawing will throw an exception here.
                        string preflightPath = PreflightRasterForPpt(resolved, -rotationDeg);

                        if (!string.IsNullOrEmpty(preflightPath))
                        {
                            safePath = preflightPath;
                        }
                    }
                    catch (System.Exception pex)
                    {
                        // Do NOT skip the image. PowerPoint handles CMYK JPGs better than .NET code.
                        // We log the warning but proceed with the original resolved path.
                        ed.WriteMessage($"\n[Warning] Preflight check failed for '{Path.GetFileName(resolved)}' (likely CMYK or locked). Attempting raw embed.");
                        ed.WriteMessage($"\nDetails: {pex.Message}");

                        // Fallback: Use original file, assume orientation is handled by PPT or post-transform
                        safePath = resolved;
                    }

                    var placement = new ImagePlacementXref
                    {
                        Path = safePath,
                        Pos = cs.Origin,
                        U = cs.Xaxis,
                        V = cs.Yaxis,
                        OriginalEntityId = img.ObjectId,
                        TargetBtrId = spaceId
                    };

                    _pendingXref.Enqueue(placement);
                    ed.WriteMessage($"\nQueued: {Path.GetFileName(safePath)}");
                }
                tr.Commit();
            }
        }

        /// <summary>
        /// Scans ACAD_IMAGE_DICT for RasterImageDefs that are actually pointing to PDF files.
        /// A RasterImageDef pointing to a PDF causes "Error 1" and crashes XREF BIND.
        /// </summary>
        private static void RemoveInvalidRasterImages(Database db, Editor ed)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!nod.Contains("ACAD_IMAGE_DICT")) return;

                    var imageDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);
                    var defsToRemove = new List<ObjectId>();
                    var keysToRemove = new List<string>();

                    ed.WriteMessage("\nChecking for invalid raster images (mismatched file types)...");

                    foreach (DBDictionaryEntry entry in imageDict)
                    {
                        var def = tr.GetObject(entry.Value, OpenMode.ForRead) as RasterImageDef;
                        if (def == null) continue;

                        string rawPath = def.SourceFileName;
                        string resolvedPath = ResolveImagePath(db, rawPath);

                        bool isInvalid = false;

                        if (string.IsNullOrEmpty(resolvedPath) || !File.Exists(resolvedPath))
                        {
                            // Optionally remove missing files too, though they usually don't crash Bind like corrupted ones do.
                            // isInvalid = true; 
                        }
                        else
                        {
                            // Check 1: Extension Mismatch (Explicitly pointing to .pdf)
                            if (Path.GetExtension(resolvedPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                            {
                                ed.WriteMessage($"\n[Detected] Image Definition pointing to PDF file: {Path.GetFileName(resolvedPath)}");
                                isInvalid = true;
                            }
                            // Check 2: Header Mismatch (File is named .png but contains PDF data)
                            else
                            {
                                try
                                {
                                    using (var fs = File.OpenRead(resolvedPath))
                                    {
                                        if (fs.Length > 4)
                                        {
                                            byte[] buffer = new byte[4];
                                            fs.Read(buffer, 0, 4);
                                            string header = System.Text.Encoding.ASCII.GetString(buffer);

                                            // Check for PDF Magic Number: %PDF
                                            if (header.StartsWith("%PDF"))
                                            {
                                                ed.WriteMessage($"\n[Detected] File is a PDF masquerading as an image: {Path.GetFileName(resolvedPath)}");
                                                isInvalid = true;
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                    // If we can't read the file, it might be corrupt anyway
                                    ed.WriteMessage($"\n[Warning] Could not read file header for: {Path.GetFileName(resolvedPath)}");
                                }
                            }
                        }

                        if (isInvalid)
                        {
                            defsToRemove.Add(def.ObjectId);
                            keysToRemove.Add(entry.Key);
                        }
                    }

                    if (defsToRemove.Count > 0)
                    {
                        // 1. Remove from Dictionary
                        foreach (string key in keysToRemove)
                        {
                            try { imageDict.Remove(key); } catch { }
                        }

                        // 2. Erase the Definition Object
                        int count = 0;
                        foreach (ObjectId id in defsToRemove)
                        {
                            try
                            {
                                var obj = tr.GetObject(id, OpenMode.ForWrite);
                                obj.Erase();
                                count++;
                            }
                            catch { }
                        }

                        // 3. (Optional) Force a Regen to clear the frames immediately
                        ed.WriteMessage($"\nRemoved {count} invalid image definition(s).");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nWarning: Error while sanitizing image dictionary: {ex.Message}");
            }
        }

        private class ImagePlacementXref
        {
            public string Path;
            public Point3d Pos;
            public Vector3d U;
            public Vector3d V;
            public ObjectId OriginalEntityId;
            public ObjectId TargetBtrId;
        }

        private static readonly Queue<ImagePlacementXref> _pendingXref = new Queue<ImagePlacementXref>();
        private static ImagePlacementXref _activePlacementXref = null;

        private static void AttachHandlersForXrefs(Database db, Document doc)
        {
            if (_handlersAttached) return; // reuse shared flags to avoid double-hook
            db.ObjectAppended += Db_ObjectAppended;
            doc.CommandWillStart += Doc_CommandWillStart_Xref;
            doc.CommandEnded += Doc_CommandEnded_Xref;
            _handlersAttached = true;
        }

        private static void DetachHandlersForXrefs(Database db, Document doc)
        {
            if (!_handlersAttached) return;
            try { db.ObjectAppended -= Db_ObjectAppended; } catch { }
            try { doc.CommandWillStart -= Doc_CommandWillStart_Xref; } catch { }
            try { doc.CommandEnded -= Doc_CommandEnded_Xref; } catch { }

            if (_pastePointHandlerAttached)
            {
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint; } catch { }
                _pastePointHandlerAttached = false;
            }

            _activePlacementXref = null;
            _activePasteDocument = null;
            _waitingForPasteStart = false;
            _handlersAttached = false;
        }

        private static void Doc_CommandWillStart_Xref(object sender, CommandEventArgs e)
        {
            if (!string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase))
                return;

            if (!_waitingForPasteStart || _activePlacementXref == null || _pastePointHandlerAttached)
                return;

            _activePasteDocument ??= AutoCADApp.DocumentManager.MdiActiveDocument;
            AutoCADApp.Idle += Application_OnIdleSendPastePoint_Xref;
            _pastePointHandlerAttached = true;
        }

        private static void Application_OnIdleSendPastePoint_Xref(object sender, EventArgs e)
        {
            try
            {
                AutoCADApp.Idle -= Application_OnIdleSendPastePoint_Xref;
                _pastePointHandlerAttached = false;

                var doc = _activePasteDocument ?? AutoCADApp.DocumentManager.MdiActiveDocument;
                var placement = _activePlacementXref;
                if (doc == null || placement == null) return;
                if (!string.Equals(doc.CommandInProgress, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;

                string x = placement.Pos.X.ToString("G17", System.Globalization.CultureInfo.InvariantCulture);
                string y = placement.Pos.Y.ToString("G17", System.Globalization.CultureInfo.InvariantCulture);
                doc.SendStringToExecute($"{x},{y}\n", true, false, false);
            }
            finally
            {
                _waitingForPasteStart = false;
            }
        }

        private static void Doc_CommandEnded_Xref(object sender, CommandEventArgs e)
        {
            if (!_isEmbeddingProcessActive ||
                !string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase))
                return;

            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            _waitingForPasteStart = false;
            if (_pastePointHandlerAttached)
            {
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint_Xref; } catch { }
                _pastePointHandlerAttached = false;
            }

            if (_pendingXref.Count == 0 && _activePlacementXref == null)
                return;

            if (_lastPastedOle.IsNull)
            {
                ed.WriteMessage("\nSkipping an item because no OLE object was created.");
                if (_pendingXref.Count > 0) _pendingXref.Dequeue();
                _activePlacementXref = null;

                if (_pendingXref.Count > 0)
                    ProcessNextPasteForXrefs(doc, ed);
                else
                    FinishEmbeddingRun(doc, ed, db);

                return;
            }

            var target = _pendingXref.Count > 0 ? _pendingXref.Dequeue() : _activePlacementXref;
            if (target == null)
            {
                ed.WriteMessage("\nError: Could not retrieve active image placement. Aborting.");
                FinishEmbeddingRun(doc, ed, db);
                return;
            }
            _activePlacementXref = null;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ole = tr.GetObject(_lastPastedOle, OpenMode.ForWrite) as Ole2Frame;

                if (ole != null)
                {
                    try
                    {
                        ApplyXrefPlacementTransform_Legacy(tr, ed, ole, target);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nFailed to transform pasted OLE: {ex.Message}");
                    }
                }

                // Erase original raster and mark its def for purge
                try
                {
                    var originalEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForWrite, false) as Entity;
                    if (originalEnt != null)
                    {
                        if (originalEnt is RasterImage imgEnt && !imgEnt.ImageDefId.IsNull)
                            _imageDefsToPurge.Add(imgEnt.ImageDefId);

                        originalEnt.Erase();
                    }
                }
                catch { }

                tr.Commit();
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            System.Threading.Thread.Sleep(50);

            if (_pendingXref.Count > 0)
            {
                ProcessNextPasteForXrefs(doc, ed);
            }
            else
            {
                _finalPastedOleForZoom = _lastPastedOle;
                FinishEmbeddingRun(doc, ed, db);
            }

            _lastPastedOle = ObjectId.Null;
        }

        /// <summary>
        /// Apply the legacy XREF-style final transform:
        /// Maps the PASTECLIP OLE extents into the oriented rectangle defined by U/V.
        /// Uses scale + translate (no extra rotation; image was pre-rotated).
        /// </summary>
        private static void ApplyXrefPlacementTransform_Legacy(
            Transaction tr,
            Editor ed,
            Ole2Frame ole,
            ImagePlacementXref target)
        {
            var ext = ole.GeometricExtents;
            Point3d oleOrigin = ext.MinPoint;
            double oleWidth = Math.Max(1e-8, ext.MaxPoint.X - ext.MinPoint.X);
            double oleHeight = Math.Max(1e-8, ext.MaxPoint.Y - ext.MinPoint.Y);

            if (oleWidth < 1e-8 || oleHeight < 1e-8)
            {
                ed.WriteMessage("\nSkipping XREF transform: invalid OLE size.");
                ole.Erase();
                return;
            }

            // Compute axis-aligned bounding box of target oriented quad
            Point3d p0 = target.Pos;
            Point3d p1 = target.Pos + target.U;
            Point3d p2 = target.Pos + target.U + target.V;
            Point3d p3 = target.Pos + target.V;

            double minX = Math.Min(Math.Min(p0.X, p1.X), Math.Min(p2.X, p3.X));
            double minY = Math.Min(Math.Min(p0.Y, p1.Y), Math.Min(p2.Y, p3.Y));
            double maxX = Math.Max(Math.Max(p0.X, p1.X), Math.Max(p2.X, p3.X));
            double maxY = Math.Max(Math.Max(p0.Y, p1.Y), Math.Max(p2.Y, p3.Y));

            double targetWidth = maxX - minX;
            double targetHeight = maxY - minY;
            if (targetWidth < 1e-8 || targetHeight < 1e-8)
            {
                ed.WriteMessage("\nSkipping XREF transform: invalid target rectangle.");
                ole.Erase();
                return;
            }

            Point3d targetMin = new Point3d(minX, minY, p0.Z);

            double scaleX = targetWidth / oleWidth;
            double scaleY = targetHeight / oleHeight;

            // Build transform: translate to origin -> scale -> translate to targetMin
            Matrix3d toOrigin = Matrix3d.Displacement(oleOrigin.GetAsVector().Negate());
            double[] scaleData =
            {
                scaleX, 0,      0, 0,
                0,      scaleY, 0, 0,
                0,      0,      1, 0,
                0,      0,      0, 1
            };
            Matrix3d scale = new Matrix3d(scaleData);
            Matrix3d toTarget = Matrix3d.Displacement(targetMin.GetAsVector());

            ole.TransformBy(toTarget * scale * toOrigin);

            try { ole.Layer = "0"; } catch { }

            // Try to move above original entity draw order (if same owner)
            try
            {
                if (!target.OriginalEntityId.IsNull)
                {
                    var origEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForRead, false) as Entity;
                    if (origEnt != null && !origEnt.IsErased)
                    {
                        var btr = tr.GetObject(ole.OwnerId, OpenMode.ForWrite) as BlockTableRecord;
                        if (btr != null && btr.DrawOrderTableId.IsValid && origEnt.OwnerId == btr.ObjectId)
                        {
                            var dot = tr.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                            if (dot != null)
                            {
                                var ids = new ObjectIdCollection(new[] { ole.ObjectId });
                                dot.MoveAbove(ids, origEnt.ObjectId);
                            }
                        }
                    }
                }
            }
            catch (System.Exception exOrder)
            {
                ed.WriteMessage($"\nWarning: Could not adjust draw order (XREF): {exOrder.Message}");
            }
        }

        /// <summary>
        /// XREF-specific paste scheduler: chooses layout, prepares clipboard, triggers PASTECLIP.
        /// </summary>
        private static void ProcessNextPasteForXrefs(Document doc, Editor ed)
        {
            if (!_isEmbeddingProcessActive || _pendingXref.Count == 0)
            {
                if (_isEmbeddingProcessActive)
                    FinishEmbeddingRun(doc, ed, doc.Database);
                return;
            }

            var target = _pendingXref.Peek();
            var db = doc.Database;
            string targetLayoutName = GetLayoutNameFromBtrId(db, target.TargetBtrId);
            if (string.IsNullOrEmpty(targetLayoutName))
            {
                _pendingXref.Dequeue();
                ProcessNextPasteForXrefs(doc, ed);
                return;
            }

            if (!PrepareClipboardWithImageShared(target.Path, ed))
            {
                _pendingXref.Dequeue();
                ProcessNextPasteForXrefs(doc, ed);
                return;
            }

            StartOleTextSizeDialogCloser(120);
            _activePlacementXref = target;
            _activePasteDocument = doc;
            _waitingForPasteStart = true;

            if (_pastePointHandlerAttached)
            {
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint_Xref; } catch { }
                _pastePointHandlerAttached = false;
            }

            string cmd = $"_.-LAYOUT S \"{targetLayoutName}\"\n_.PASTECLIP\n";
            ed.WriteMessage($"\nActivating layout '{targetLayoutName}' for pasting...");
            doc.SendStringToExecute(cmd, true, false, false);
        }

        /// <summary>
        /// Prepare clipboard for XREF: uses shared PPT helper but only needs path.
        /// This wraps the shared ImagePlacement-based method for compatibility.
        /// </summary>
        private static bool PrepareClipboardWithImageShared(string imagePath, Editor ed)
        {
            var placement = new ImagePlacement
            {
                Path = imagePath
            };
            return PrepareClipboardWithImageShared(placement, ed);
        }


        /// <summary>
        /// XREF-specific attempt to find title block extents for zoom (unchanged).
        /// </summary>
        private static bool TryGetTitleBlockOutlinePointsForEmbed(Database db, out Point3d[] poly)
        {
            poly = null;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lm = LayoutManager.Current;
                    if (lm == null) return false;

                    var layId = lm.GetLayoutId(lm.CurrentLayout);
                    var layout = (Layout)tr.GetObject(layId, OpenMode.ForRead);
                    var psBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                    string[] hints = { "x-tb", "title", "tblock", "border", "sheet" };
                    BlockReference best = null;
                    double bestArea = 0;

                    foreach (ObjectId id in psBtr)
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;

                        string name = null;
                        try
                        {
                            if (br.IsDynamicBlock)
                            {
                                var dyn = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                                name = dyn?.Name;
                            }
                            else
                            {
                                var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                                name = btr?.Name;
                            }
                        }
                        catch { }

                        if (string.IsNullOrEmpty(name)) continue;

                        var lname = name.ToLowerInvariant();
                        bool matches = false;
                        foreach (var h in hints)
                        {
                            if (lname.Contains(h))
                            {
                                matches = true;
                                break;
                            }
                        }
                        if (!matches) continue;

                        Extents3d ext;
                        try { ext = br.GeometricExtents; }
                        catch { continue; }

                        double area = Math.Abs((ext.MaxPoint.X - ext.MinPoint.X) *
                                               (ext.MaxPoint.Y - ext.MinPoint.Y));
                        if (area > bestArea)
                        {
                            best = br;
                            bestArea = area;
                        }
                    }

                    if (best != null)
                    {
                        var ex = best.GeometricExtents;
                        var pmin = ex.MinPoint;
                        var pmax = ex.MaxPoint;
                        poly = new[]
                        {
                            new Point3d(pmin.X, pmin.Y, 0),
                            new Point3d(pmax.X, pmin.Y, 0),
                            new Point3d(pmax.X, pmax.Y, 0),
                            new Point3d(pmin.X, pmax.Y, 0)
                        };
                        return true;
                    }

                    var dbMin = db.Pextmin;
                    var dbMax = db.Pextmax;
                    poly = new[]
                    {
                        new Point3d(dbMin.X, dbMin.Y, 0),
                        new Point3d(dbMax.X, dbMin.Y, 0),
                        new Point3d(dbMax.X, dbMax.Y, 0),
                        new Point3d(dbMin.X, dbMax.Y, 0)
                    };
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void ZoomToTitleBlockForEmbed(Editor ed, Point3d[] poly)
        {
            if (ed == null || poly == null || poly.Length < 2) return;

            try
            {
                var ext = new Extents3d(poly[0], poly[1]);
                for (int i = 2; i < poly.Length; i++)
                    ext.AddPoint(poly[i]);

                double margin = Math.Max(ext.MaxPoint.X - ext.MinPoint.X,
                                         ext.MaxPoint.Y - ext.MinPoint.Y) * 0.05;

                var pMin = new Point3d(ext.MinPoint.X - margin, ext.MinPoint.Y - margin, 0);
                var pMax = new Point3d(ext.MaxPoint.X + margin, ext.MaxPoint.Y + margin, 0);

                using (var view = ed.GetCurrentView())
                {
                    view.Width = pMax.X - pMin.X;
                    view.Height = pMax.Y - pMin.Y;
                    view.CenterPoint =
                        new Point2d((pMin.X + pMax.X) / 2.0, (pMin.Y + pMax.Y) / 2.0);
                    ed.SetCurrentView(view);
                }
                ed.Regen();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during zoom: {ex.Message}");
            }
        }

        private static Point2d? GetPdfPageSizeInches(string pdfFilePath, int pageNumber)
        {
            if (string.IsNullOrEmpty(pdfFilePath) || !File.Exists(pdfFilePath) || pageNumber <= 0)
            {
                return null;
            }

            try
            {
                using (var pdf = new PdfDocument())
                {
                    pdf.LoadFromFile(pdfFilePath);

                    if (pageNumber > pdf.Pages.Count)
                    {
                        return null; // Page number is out of range
                    }

                    var page = pdf.Pages[pageNumber - 1]; // Spire.Pdf is 0-indexed
                    var sizeInPoints = page.Size;

                    var unitConverter = new PdfUnitConvertor();
                    float widthInches = unitConverter.ConvertUnits(sizeInPoints.Width, PdfGraphicsUnit.Point, PdfGraphicsUnit.Inch);
                    float heightInches = unitConverter.ConvertUnits(sizeInPoints.Height, PdfGraphicsUnit.Point, PdfGraphicsUnit.Inch);

                    return new Point2d(widthInches, heightInches);
                }
            }
            catch (System.Exception)
            {
                // Log exception here if needed
                return null;
            }
        }
    }
}
