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
using System.Runtime.InteropServices; // Recommended for releasing COM objects
namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
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
                            if (id.ObjectClass.DxfName.ToUpperInvariant() == "INSERT")
                            {
                                var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                                if (br != null)
                                {
                                    blockReferences.Add(br);
                                }
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
                                    var ent = obj as Entity;
                                    if (ent != null)
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
        /// Finds a likely title block reference in the Model Space and explodes it.
        /// This is used to "unpack" a composite title block so that nested images can be found for embedding.
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
                    // Target the Model Space for the search.
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    string[] hints = { "title sheet", "_wm_border", "border", "tblock", "title", "sheet", "x-tb" };
                    BlockReference bestCandidate = null;
                    string matchedName = "";
                    double maxArea = 0;

                    // Find the largest block reference in Model Space that looks like a title block.
                    foreach (ObjectId id in modelSpace)
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;

                        var potentialNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        // Defensively gather all possible names associated with the block reference.
                        // 1. Get name from main BlockTableRecord.
                        try
                        {
                            var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            if (btr != null)
                            {
                                if (!string.IsNullOrEmpty(btr.Name))
                                {
                                    potentialNames.Add(btr.Name);
                                }
                                // 2. Get XREF filename if applicable.
                                if (btr.IsFromExternalReference && !string.IsNullOrEmpty(btr.PathName))
                                {
                                    potentialNames.Add(Path.GetFileNameWithoutExtension(btr.PathName));
                                }
                            }
                        }
                        catch { /* Ignore errors reading primary BTR */ }

                        // 3. Get name from Dynamic BlockTableRecord if it's a dynamic block.
                        if (br.IsDynamicBlock)
                        {
                            try
                            {
                                var dynBtr = tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (dynBtr != null && !string.IsNullOrEmpty(dynBtr.Name))
                                {
                                    potentialNames.Add(dynBtr.Name);
                                }
                            }
                            catch { /* Ignore errors reading dynamic BTR */ }
                        }

                        // Now, check all gathered names against the hints.
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

                        if (isMatch)
                        {
                            try
                            {
                                var ext = br.GeometricExtents;
                                double area = (ext.MaxPoint.X - ext.MinPoint.X) * (ext.MaxPoint.Y - ext.MinPoint.Y);
                                if (area > maxArea)
                                {
                                    maxArea = area;
                                    bestCandidate = br;
                                    matchedName = currentMatch;
                                }
                            }
                            catch { /* Ignore blocks that fail to get extents */ }
                        }
                    }

                    if (bestCandidate != null)
                    {
                        ed.WriteMessage($"\nFound title block reference in Model Space matching name '{matchedName}'. Exploding it...");

                        var explodedEntities = new DBObjectCollection();
                        bestCandidate.UpgradeOpen();
                        bestCandidate.Explode(explodedEntities);

                        foreach (DBObject obj in explodedEntities)
                        {
                            var ent = obj as Entity;
                            if (ent != null)
                            {
                                // Add the exploded entities to Model Space.
                                modelSpace.AppendEntity(ent);
                                tr.AddNewlyCreatedDBObject(ent, true);
                            }
                        }

                        bestCandidate.Erase();
                        ed.WriteMessage("\nExplode complete.");
                    }
                    else
                    {
                        ed.WriteMessage("\nNo composite title block reference found to explode in Model Space.");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError while trying to find and explode title block: {ex.Message}");
            }
        }


        [CommandMethod("EMBEDFROMXREFS", CommandFlags.Modal)]
        public static void EmbedFromXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage("\n--- Starting EMBEDFROMXREFS ---");

            _pending.Clear();
            _lastPastedOle = ObjectId.Null;
            _isEmbeddingProcessActive = false;
            _finalPastedOleForZoom = ObjectId.Null;
            _xrefsToDetach.Clear();

            try
            {
                DeleteOldEmbedTemps(daysOld: 7);
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

                // Phase 2: scan and queue images
                try
                {
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

                    ed.WriteMessage("\nScanning all layouts for raster images...");
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var modelSpaceId = bt[BlockTableRecord.ModelSpace];

                        ed.WriteMessage("\n- Scanning Model Space...");
                        CollectAndPreflightImagesInLayout(doc, modelSpaceId);

                        var layoutDictId = db.LayoutDictionaryId;
                        var layoutDict = (DBDictionary)tr.GetObject(layoutDictId, OpenMode.ForRead);

                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            if (entry.Key.Equals("Model", StringComparison.OrdinalIgnoreCase))
                                continue;

                            var layoutId = entry.Value;
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            ed.WriteMessage($"\n- Scanning layout: {layout.LayoutName}...");
                            CollectAndPreflightImagesInLayout(doc, layout.BlockTableRecordId);
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[!] A critical error occurred during image collection: {ex.Message}");
                    return;
                }

                if (_pending.Count == 0)
                {
                    ed.WriteMessage("\nNo valid raster images found to embed.");
                    _isEmbeddingProcessActive = true;
                    FinishEmbeddingRun(doc, ed, db);
                    return;
                }

                ed.WriteMessage($"\nSuccessfully queued {_pending.Count} image(s).");

                if (!EnsurePowerPoint(ed))
                {
                    ed.WriteMessage("\n[!] Failed to start or connect to PowerPoint. Aborting.");
                    return;
                }

                if (WindowOrchestrator.TryGetPowerPointHwnd(out var pptHwnd))
                {
                    WindowOrchestrator.EnsureSeparationOrSafeOverlap(ed, pptHwnd, preferDifferentMonitor: true);
                }

                AttachHandlers(db, doc);
                ed.WriteMessage($"\nStarting paste process for {_pending.Count} raster image(s)...");
                _isEmbeddingProcessActive = true;

                ProcessNextPaste(doc, ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[!] An error occurred during the PowerPoint embedding process: {ex.Message}");
            }
            finally
            {
                if (!_isEmbeddingProcessActive && _pending.Count == 0)
                {
                    ClosePowerPoint(ed);
                    RestoreOriginalLayer(db, originalClayer);
                }
                else if (!_isEmbeddingProcessActive && _pending.Count > 0)
                {
                    ed.WriteMessage("\nAborting embedding process due to an earlier error.");
                    DetachHandlers(db, doc);
                    ClosePowerPoint(ed);
                    RestoreOriginalLayer(db, originalClayer);
                    _pending.Clear();
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

        private static void CollectAndPreflightImagesInLayout(Document doc, ObjectId spaceId)
        {
            var db = doc.Database;
            var ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var space = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);

                foreach (ObjectId id in space)
                {
                    if (id.ObjectClass.DxfName.ToUpperInvariant() != "IMAGE") continue;

                    var img = tr.GetObject(id, OpenMode.ForRead) as RasterImage;
                    if (img == null || img.ImageDefId.IsNull) continue;

                    var def = tr.GetObject(img.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                    if (def == null) continue;

                    var ownerBtr = tr.GetObject(img.OwnerId, OpenMode.ForRead) as BlockTableRecord;
                    if (ownerBtr != null && ownerBtr.IsFromExternalReference)
                        _xrefsToDetach.Add(ownerBtr.ObjectId);

                    string resolved = ResolveImagePath(db, def.SourceFileName);
                    if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                    {
                        ed.WriteMessage($"\nSkipping missing image: {def.SourceFileName}");
                        continue;
                    }

                    string safePath = null;
                    var cs = img.Orientation;
                    try
                    {
                        ed.WriteMessage($"\nProcessing: {Path.GetFileName(resolved)}...");

                        double rotationInDegrees =
                            Math.Atan2(cs.Xaxis.Y, cs.Xaxis.X) * (180.0 / Math.PI);

                        safePath = PreflightRasterForPpt(resolved, -rotationInDegrees);
                    }
                    catch (System.Exception pex)
                    {
                        ed.WriteMessage($"\n[!] PREFLIGHT FAILED for '{Path.GetFileName(resolved)}': {pex.Message} — skipping this image.");
                        continue;
                    }

                    if (string.IsNullOrEmpty(safePath))
                    {
                        ed.WriteMessage($"\n[!] Preflight returned an empty path for '{Path.GetFileName(resolved)}'. Skipping.");
                        continue;
                    }

                    var placement = new ImagePlacement
                    {
                        Path = safePath,
                        Pos = cs.Origin,
                        U = cs.Xaxis,
                        V = cs.Yaxis,
                        OriginalEntityId = img.ObjectId,
                        TargetBtrId = spaceId,
                        Source = PlacementSource.Xref
                    };

                    _pending.Enqueue(placement);
                    ed.WriteMessage($"\nQueued [{_pending.Count}]: {Path.GetFileName(placement.Path)}");
                }
                tr.Commit();
            }
        }

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

                    var psId = layout.BlockTableRecordId;
                    var psBtr = (BlockTableRecord)tr.GetObject(psId, OpenMode.ForRead);

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
                        foreach (var h in hints) { if (lname.Contains(h)) { matches = true; break; } }
                        if (!matches) continue;

                        Extents3d ext;
                        try { ext = br.GeometricExtents; } catch { continue; }

                        double area = Math.Abs((ext.MaxPoint.X - ext.MinPoint.X) * (ext.MaxPoint.Y - ext.MinPoint.Y));
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
                        new Point3d(pmin.X, pmin.Y, 0), new Point3d(pmax.X, pmin.Y, 0),
                        new Point3d(pmax.X, pmax.Y, 0), new Point3d(pmin.X, pmax.Y, 0)
                    };
                        return true;
                    }

                    var pMin = db.Pextmin;
                    var pMax = db.Pextmax;
                    poly = new[]
                    {
                    new Point3d(pMin.X, pMin.Y, 0), new Point3d(pMax.X, pMin.Y, 0),
                    new Point3d(pMax.X, pMax.Y, 0), new Point3d(pMin.X, pMax.Y, 0)
                };
                    return true;
                }
            }
            catch { return false; }
        }

        private static void ZoomToTitleBlockForEmbed(Editor ed, Point3d[] poly)
        {
            if (ed == null || poly == null || poly.Length < 2) return;
            try
            {
                var ext = new Extents3d(poly[0], poly[1]);
                for (int i = 2; i < poly.Length; i++) ext.AddPoint(poly[i]);

                double margin = Math.Max(ext.MaxPoint.X - ext.MinPoint.X, ext.MaxPoint.Y - ext.MinPoint.Y) * 0.05;

                Point3d pMin = new Point3d(ext.MinPoint.X - margin, ext.MinPoint.Y - margin, 0);
                Point3d pMax = new Point3d(ext.MaxPoint.X + margin, ext.MaxPoint.Y + margin, 0);

                using (var view = ed.GetCurrentView())
                {
                    view.Width = pMax.X - pMin.X;
                    view.Height = pMax.Y - pMin.Y;
                    view.CenterPoint = new Point2d((pMin.X + pMax.X) / 2.0, (pMin.Y + pMax.Y) / 2.0);
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