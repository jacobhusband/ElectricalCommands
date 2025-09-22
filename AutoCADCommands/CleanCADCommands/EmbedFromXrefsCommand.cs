using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("EMBEDFROMXREFS", CommandFlags.Modal)]
        public static void EmbedFromXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            // NEW: center the view on the title block (non-fatal if not found)
            try
            {
                if (TryGetTitleBlockOutlinePointsForEmbed(db, out var tbPoly) && tbPoly != null && tbPoly.Length > 0)
                {
                    ZoomToTitleBlockForEmbed(ed, tbPoly);
                }
            }
            catch { /* ignore; zoom is best-effort */ }

            _pending.Clear();
            _lastPastedOle = ObjectId.Null;

            try
            {
                // Ensure layer "0" is thawed and make it current for embedding
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has("0"))
                    {
                        var zeroId = lt["0"];
                        var zeroLtr = (LayerTableRecord)tr.GetObject(zeroId, OpenMode.ForWrite);
                        if (zeroLtr.IsFrozen) zeroLtr.IsFrozen = false;
                        if (_savedClayer.IsNull) _savedClayer = db.Clayer;
                        db.Clayer = zeroId;
                    }
                    tr.Commit();
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id in space)
                    {
                        var img = tr.GetObject(id, OpenMode.ForRead) as RasterImage;
                        if (img == null) continue;
                        if (img.ImageDefId.IsNull) continue;

                        var def = tr.GetObject(img.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                        if (def == null) continue;

                        string resolved = ResolveImagePath(db, def.SourceFileName);
                        if (string.IsNullOrWhiteSpace(resolved))
                        {
                            ed.WriteMessage($"\nSkipping missing image: {def.SourceFileName}");
                            continue;
                        }

                        var cs = img.Orientation;
                        var placement = new ImagePlacement
                        {
                            Path = resolved,
                            Pos = cs.Origin,
                            U = cs.Xaxis,
                            V = cs.Yaxis,
                            ImageId = img.ObjectId
                        };
                        _pending.Enqueue(placement);
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to collect raster images: {ex.Message}");
                // Restore original current layer if we changed it
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                return;
            }

            if (_pending.Count == 0)
            {
                ed.WriteMessage("\nNo raster images found in current space.");
                // Close PowerPoint if it was open from a prior run (no save)
                ClosePowerPoint(ed);
                // Restore original current layer if we changed it
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                // If running as part of CLEANCAD chain, move on to FINALIZE immediately
                if (_chainFinalizeAfterEmbed)
                {
                    _chainFinalizeAfterEmbed = false;
                    doc.SendStringToExecute("_.FINALIZE ", true, false, false);
                }
                return;
            }

            if (!EnsurePowerPoint(ed))
            {
                // Restore original current layer if we changed it
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                return;
            }
            AttachHandlers(db, doc);
            ed.WriteMessage($"\nEmbedding over {_pending.Count} raster image(s)...");
            ProcessNextPaste(doc, ed);

        }

        // --- helpers (unique names to avoid collisions) ---

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

                    // common name hints; pick largest match
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
                                name = br.Name;
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
                            new Point3d(pmin.X, pmin.Y, 0),
                            new Point3d(pmax.X, pmin.Y, 0),
                            new Point3d(pmax.X, pmax.Y, 0),
                            new Point3d(pmin.X, pmax.Y, 0)
                        };
                        return true;
                    }

                    // fallback: paper extents
                    var pMin = db.Pextmin;
                    var pMax = db.Pextmax;
                    poly = new[]
                    {
                        new Point3d(pMin.X, pMin.Y, 0),
                        new Point3d(pMax.X, pMin.Y, 0),
                        new Point3d(pMax.X, pMax.Y, 0),
                        new Point3d(pMin.X, pMax.Y, 0)
                    };
                    return true;
                }
            }
            catch { return false; }
        }

        private static void ZoomToTitleBlockForEmbed(Editor ed, Point3d[] poly)
        {
            if (ed == null || poly == null || poly.Length == 0) return;

            try
            {
                var ext = new Extents3d(poly[0], poly[0]);
                for (int i = 1; i < poly.Length; i++) ext.AddPoint(poly[i]);

                double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
                double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);
                double margin = Math.Max(width, height) * 0.05;

                Point3d pMin = new Point3d(ext.MinPoint.X - margin, ext.MinPoint.Y - margin, ext.MinPoint.Z);
                Point3d pMax = new Point3d(ext.MaxPoint.X + margin, ext.MaxPoint.Y + margin, ext.MaxPoint.Z);

                CleanupCommands.Zoom(pMin, pMax, Point3d.Origin, 1.0);
                try { ed.Regen(); } catch { }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during zoom: {ex.Message}");
            }
        }
    }
}
