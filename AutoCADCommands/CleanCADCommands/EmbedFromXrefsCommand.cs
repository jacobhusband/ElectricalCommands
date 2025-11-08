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
        /// Try to find and explode a composite title block (unchanged helper).
        /// </summary>
        public static void FindAndExplodeTitleBlockReference()
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

                    string[] hints = { "title sheet", "_wm_border", "border", "tblock", "title", "sheet", "x-tb" };
                    BlockReference bestCandidate = null;
                    string matchedName = "";
                    double maxArea = 0;

                    foreach (ObjectId id in modelSpace)
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;

                        var potentialNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        try
                        {
                            var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            if (btr != null)
                            {
                                if (!string.IsNullOrEmpty(btr.Name))
                                    potentialNames.Add(btr.Name);

                                if (btr.IsFromExternalReference && !string.IsNullOrEmpty(btr.PathName))
                                    potentialNames.Add(Path.GetFileNameWithoutExtension(btr.PathName));
                            }
                        }
                        catch { }

                        if (br.IsDynamicBlock)
                        {
                            try
                            {
                                var dynBtr = tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (dynBtr != null && !string.IsNullOrEmpty(dynBtr.Name))
                                    potentialNames.Add(dynBtr.Name);
                            }
                            catch { }
                        }

                        bool isMatch = false;
                        string currentMatch = "";
                        foreach (var name in potentialNames)
                        {
                            foreach (var hint in hints)
                            {
                                if (name.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    isMatch = true;
                                    currentMatch = name;
                                    break;
                                }
                            }
                            if (isMatch) break;
                        }

                        if (!isMatch) continue;

                        try
                        {
                            var ext = br.GeometricExtents;
                            double area = (ext.MaxPoint.X - ext.MinPoint.X) *
                                          (ext.MaxPoint.Y - ext.MinPoint.Y);
                            if (area > maxArea)
                            {
                                maxArea = area;
                                bestCandidate = br;
                                matchedName = currentMatch;
                            }
                        }
                        catch { }
                    }

                    if (bestCandidate != null)
                    {
                        ed.WriteMessage($"\nFound title block reference matching '{matchedName}'. Exploding it...");
                        var exploded = new DBObjectCollection();
                        bestCandidate.UpgradeOpen();
                        bestCandidate.Explode(exploded);

                        foreach (DBObject obj in exploded)
                        {
                            if (obj is Entity ent)
                            {
                                modelSpace.AppendEntity(ent);
                                tr.AddNewlyCreatedDBObject(ent, true);
                            }
                        }

                        bestCandidate.Erase();
                        ed.WriteMessage("\nExplode complete.");
                    }
                    else
                    {
                        ed.WriteMessage("\nNo composite title block reference found in Model Space.");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError while trying to find and explode title block: {ex.Message}");
            }
        }

        /// <summary>
        /// EMBEDFROMXREFS:
        /// - Collects all raster images (including from XREFs) across layouts.
        /// - Pre-rotates each image to match its orientation.
        /// - Uses PowerPoint as an OLE source and pastes into the correct layout.
        /// - Transforms the pasted OLE to align with the original image's oriented rectangle.
        /// - Records XREFs for detaching and image defs for purging.
        /// This implementation mirrors your previously working rotated-raster behavior.
        /// </summary>
        [CommandMethod("EMBEDFROMXREFS", CommandFlags.Modal)]
        public static void EmbedFromXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage("\n--- Starting EMBEDFROMXREFS ---");

            // Reset shared state used by the XREF pipeline.
            _pendingXref.Clear();
            _lastPastedOle = ObjectId.Null;
            _isEmbeddingProcessActive = false;
            _finalPastedOleForZoom = ObjectId.Null;
            _xrefsToDetach.Clear();

            try
            {
                DeleteOldEmbedTemps(7);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Warning] Could not delete old temp files: {ex.Message}");
            }

            ObjectId originalClayer = db.Clayer;
            Matrix3d originalUcs = ed.CurrentUserCoordinateSystem;
            string originalLayout = LayoutManager.Current.CurrentLayout;

            try
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;

                // Ensure layer "0" is thawed and set active
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has("0"))
                    {
                        var zeroId = lt["0"];
                        var zeroLtr = (LayerTableRecord)tr.GetObject(zeroId, OpenMode.ForWrite);
                        if (zeroLtr.IsFrozen)
                        {
                            ed.WriteMessage("\nThawing layer '0'.");
                            zeroLtr.IsFrozen = false;
                        }
                        if (_savedClayer.IsNull) _savedClayer = originalClayer;
                        db.Clayer = zeroId;
                    }
                    tr.Commit();
                }

                // Collect images from all layouts
                try
                {
                    ed.WriteMessage("\nScanning all layouts for raster images...");
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var modelSpaceId = bt[BlockTableRecord.ModelSpace];

                        ed.WriteMessage("\n- Scanning Model Space...");
                        CollectAndPreflightImagesInLayoutForXrefs(doc, modelSpaceId);

                        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            if (entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
                                continue;

                            var layoutId = entry.Value;
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            ed.WriteMessage($"\n- Scanning layout: {layout.LayoutName}...");
                            CollectAndPreflightImagesInLayoutForXrefs(doc, layout.BlockTableRecordId);
                        }

                        tr.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[!] Critical error during image collection: {ex.Message}");
                    return;
                }

                if (_pendingXref.Count == 0)
                {
                    ed.WriteMessage("\nNo valid raster images found to embed.");
                    _isEmbeddingProcessActive = true;
                    FinishEmbeddingRun(doc, ed, db);
                    return;
                }

                ed.WriteMessage($"\nSuccessfully queued {_pendingXref.Count} raster image(s).");

                if (!EnsurePowerPoint(ed))
                {
                    ed.WriteMessage("\n[!] Failed to start or connect to PowerPoint. Aborting.");
                    return;
                }

                if (WindowOrchestrator.TryGetPowerPointHwnd(out var pptHwnd))
                {
                    WindowOrchestrator.EnsureSeparationOrSafeOverlap(ed, pptHwnd, preferDifferentMonitor: true);
                }

                AttachHandlersForXrefs(db, doc);
                _isEmbeddingProcessActive = true;

                ed.WriteMessage($"\nStarting paste process for {_pendingXref.Count} raster image(s)...");
                ProcessNextPasteForXrefs(doc, ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[!] Error during EMBEDFROMXREFS: {ex.Message}");
            }
            finally
            {
                if (!_isEmbeddingProcessActive && _pendingXref.Count == 0)
                {
                    ClosePowerPoint(ed);
                    RestoreOriginalLayer(db, originalClayer);
                }
                else if (!_isEmbeddingProcessActive && _pendingXref.Count > 0)
                {
                    ed.WriteMessage("\nAborting embedding process due to an earlier error.");
                    DetachHandlersForXrefs(db, doc);
                    ClosePowerPoint(ed);
                    RestoreOriginalLayer(db, originalClayer);
                    _pendingXref.Clear();
                }

                ed.CurrentUserCoordinateSystem = originalUcs;

                try
                {
                    if (LayoutManager.Current.CurrentLayout != originalLayout)
                        LayoutManager.Current.CurrentLayout = originalLayout;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Could not restore original layout: {ex.Message}");
                }
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
                    if (!id.ObjectClass.DxfName.Equals("IMAGE", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var img = tr.GetObject(id, OpenMode.ForRead) as RasterImage;
                    if (img == null || img.ImageDefId.IsNull) continue;

                    var def = tr.GetObject(img.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                    if (def == null) continue;

                    // If hosted in an XREF, mark that XREF for detachment later.
                    if (tr.GetObject(img.OwnerId, OpenMode.ForRead) is BlockTableRecord ownerBtr &&
                        ownerBtr.IsFromExternalReference)
                    {
                        _xrefsToDetach.Add(ownerBtr.ObjectId);
                    }

                    string resolved = ResolveImagePath(db, def.SourceFileName);
                    if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                    {
                        ed.WriteMessage($"\nSkipping missing image: {def.SourceFileName}");
                        continue;
                    }

                    string safePath;
                    var cs = img.Orientation;
                    try
                    {
                        ed.WriteMessage($"\nProcessing: {Path.GetFileName(resolved)}...");

                        // AutoCAD orientation: counter-clockwise; GDI+ RotateTransform: clockwise.
                        double rotationDeg = Math.Atan2(cs.Xaxis.Y, cs.Xaxis.X) * (180.0 / Math.PI);
                        safePath = PreflightRasterForPpt(resolved, -rotationDeg);
                    }
                    catch (System.Exception pex)
                    {
                        ed.WriteMessage($"\n[!] PREFLIGHT FAILED for '{Path.GetFileName(resolved)}': {pex.Message} — skipping.");
                        continue;
                    }

                    if (string.IsNullOrEmpty(safePath))
                    {
                        ed.WriteMessage($"\n[!] Preflight returned empty path for '{Path.GetFileName(resolved)}'. Skipping.");
                        continue;
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
                    ed.WriteMessage($"\nQueued [{_pendingXref.Count}]: {Path.GetFileName(placement.Path)}");
                }
                tr.Commit();
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
    }
}