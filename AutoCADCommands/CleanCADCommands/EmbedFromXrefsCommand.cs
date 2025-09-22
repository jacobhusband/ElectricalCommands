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

            // try to focus the title block area (best-effort)
            try
            {
                if (TryGetTitleBlockOutlinePointsForEmbed(db, out var tbPoly) && tbPoly != null && tbPoly.Length > 0)
                    ZoomToTitleBlockForEmbed(ed, tbPoly);
            }
            catch { /* ignore */ }

            _pending.Clear();
            _lastPastedOle = ObjectId.Null;

            // NEW: purge old temp files from previous runs (best effort)
            try { DeleteOldEmbedTemps(daysOld: 7); } catch { }

            // --- collect raster images from current space ---
            int queued = 0;
            try
            {
                // Ensure layer "0" is thawed and current
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

                    int index = 0;
                    foreach (ObjectId id in space)
                    {
                        var img = tr.GetObject(id, OpenMode.ForRead) as RasterImage;
                        if (img == null) continue;
                        if (img.ImageDefId.IsNull) continue;

                        var def = tr.GetObject(img.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                        if (def == null) continue;

                        string resolved = ResolveImagePath(db, def.SourceFileName);
                        if (string.IsNullOrWhiteSpace(resolved) || !File.Exists(resolved))
                        {
                            ed.WriteMessage($"\nSkipping missing image: {def?.SourceFileName}");
                            continue;
                        }

                        // NEW: preflight & sanitize the raster to something PPT/clipboard will reliably accept
                        string safePath = null;
                        try
                        {
                            safePath = PreflightRasterForPpt(resolved);
                        }
                        catch (System.Exception pex)
                        {
                            ed.WriteMessage($"\n[!] Preflight failed for “{Path.GetFileName(resolved)}”: {pex.Message} — skipping this image.");
                            continue; // isolation: do not crash the whole run on a single bad file
                        }

                        var cs = img.Orientation;
                        var placement = new ImagePlacement
                        {
                            Path = safePath ?? resolved, // fall back to original if preflight returned null
                            Pos = cs.Origin,
                            U = cs.Xaxis,
                            V = cs.Yaxis,
                            ImageId = img.ObjectId,
                            // Optional: only if your type has this property/field
                            // Index   = ++index
                        };

                        _pending.Enqueue(placement);
                        queued++;
                        ed.WriteMessage($"\nQueued [{++index}]: {Path.GetFileName(placement.Path)}");
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to collect raster images: {ex.Message}");
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                return;
            }

            if (queued == 0)
            {
                ed.WriteMessage("\nNo raster images found in current space.");
                ClosePowerPoint(ed);
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;

                if (_chainFinalizeAfterEmbed)
                {
                    _chainFinalizeAfterEmbed = false;
                    doc.SendStringToExecute("_.FINALIZE ", true, false, false);
                }
                return;
            }

            if (!EnsurePowerPoint(ed))
            {
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                return;
            }

            if (WindowOrchestrator.TryGetPowerPointHwnd(out var pptHwnd /*, pptAppHwndHint: new IntPtr(pptApp.HWND)*/))
            {
                WindowOrchestrator.EnsureSeparationOrSafeOverlap(ed, pptHwnd, preferDifferentMonitor: true);
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

        /// <summary>
        /// Load the raster safely for PPT/clipboard: flatten multipage TIFFs, convert indexed/CMYK/1bpp
        /// to 24bpp RGB, and optionally downscale very large images. Returns a temp file path (PNG).
        /// </summary>
        private static string PreflightRasterForPpt(string srcPath)
        {
            // Always normalize to PNG in %TEMP%\AutoCADCleanupTool\embed\
            string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "embed");
            Directory.CreateDirectory(tempDir);
            string outPath = Path.Combine(
                tempDir,
                Path.GetFileNameWithoutExtension(srcPath) + "_ppt.png"
            );

            using (var orig = System.Drawing.Image.FromFile(srcPath))
            {
                // pick first frame if multi-page (e.g., TIFF)
                try
                {
                    var fd = new FrameDimension(orig.FrameDimensionsList[0]);
                    if (orig.GetFrameCount(fd) > 1)
                        orig.SelectActiveFrame(fd, 0);
                }
                catch { /* not multi-frame; ignore */ }

                int maxSide = 8000; // safety cap to avoid huge clipboard COM allocations

                int w = orig.Width, h = orig.Height;

                // Compute downscale if necessary
                double scale = 1.0;
                if (Math.Max(w, h) > maxSide)
                    scale = (double)maxSide / Math.Max(w, h);

                int targetW = Math.Max(1, (int)Math.Round(w * scale));
                int targetH = Math.Max(1, (int)Math.Round(h * scale));

                // Render to 24bpp RGB bitmap
                using (var bmp = new Bitmap(targetW, targetH, PixelFormat.Format24bppRgb))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(System.Drawing.Color.White);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.DrawImage(orig, new Rectangle(0, 0, targetW, targetH));

                    // PNG encoder (lossless, safe for linework/scans)
                    bmp.Save(outPath, ImageFormat.Png);
                }
            }

            return outPath;
        }

        /// <summary>
        /// Clean up stale temp files to keep %TEMP% tidy.
        /// </summary>
        private static void DeleteOldEmbedTemps(int daysOld = 7)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "embed");
            if (!Directory.Exists(tempDir)) return;

            var cutoff = DateTime.Now.AddDays(-Math.Abs(daysOld));
            foreach (var f in Directory.EnumerateFiles(tempDir, "*.png"))
            {
                try
                {
                    var info = new FileInfo(f);
                    if (info.LastWriteTime < cutoff) info.Delete();
                }
                catch { /* ignore */ }
            }
        }
    }
}
