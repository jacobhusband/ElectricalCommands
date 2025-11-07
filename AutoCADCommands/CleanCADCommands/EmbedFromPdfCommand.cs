using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using ImageMagick;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("EMBEDFROMPDFS", CommandFlags.Modal)]
        public static void EmbedFromPdfs()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage("\n--- Starting EMBEDFROMPDFS ---");

            _pending.Clear();
            _lastPastedOle = ObjectId.Null;
            _isEmbeddingProcessActive = false;
            _finalPastedOleForZoom = ObjectId.Null;

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

            try
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;
                _savedClayer = originalClayer;

                try
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                        if (lt.Has("0"))
                        {
                            var zeroId = lt["0"];
                            if (db.Clayer != zeroId)
                            {
                                db.Clayer = zeroId;
                            }
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[Warning] Could not switch to layer '0': {ex.Message}");
                }

                int queuedCount = CollectAndQueuePdfUnderlays(doc);
                if (queuedCount == 0)
                {
                    ed.WriteMessage("\nNo PDF underlays found to embed.");
                    RestoreOriginalLayer(db, originalClayer);
                    return;
                }

                ed.WriteMessage($"\nQueued {queuedCount} PDF underlay(s) for embedding.");

                if (!EnsurePowerPoint(ed))
                {
                    ed.WriteMessage("\n[!] Failed to start or connect to PowerPoint. Aborting.");
                    RestoreOriginalLayer(db, originalClayer);
                    return;
                }

                if (WindowOrchestrator.TryGetPowerPointHwnd(out var pptHwnd))
                {
                    WindowOrchestrator.EnsureSeparationOrSafeOverlap(ed, pptHwnd, preferDifferentMonitor: true);
                }

                AttachHandlers(db, doc);
                _isEmbeddingProcessActive = true;
                ProcessNextPaste(doc, ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[!] An error occurred during PDF embedding: {ex.Message}");
            }
            finally
            {
                ed.CurrentUserCoordinateSystem = originalUcs;

                if (!_isEmbeddingProcessActive)
                {
                    DetachHandlers(db, doc);
                    ClosePowerPoint(ed);
                    RestoreOriginalLayer(db, originalClayer);
                    _pending.Clear();
                }
            }
        }

        private static int CollectAndQueuePdfUnderlays(Document doc)
        {
            var db = doc.Database;
            var ed = doc.Editor;
            int queued = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                    if (btr == null || !btr.IsLayout)
                        continue;

                    foreach (ObjectId entId in btr)
                    {
                        if (!entId.ObjectClass.DxfName.Equals("PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                            continue;

                        var pdfRef = tr.GetObject(entId, OpenMode.ForRead) as UnderlayReference;
                        if (pdfRef == null)
                            continue;

                        var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                        if (pdfDef == null)
                            continue;

                        string resolvedPdfPath = ResolveImagePath(db, pdfDef.SourceFileName);
                        if (string.IsNullOrEmpty(resolvedPdfPath) || !File.Exists(resolvedPdfPath))
                        {
                            ed.WriteMessage($"\nSkipping missing PDF file: {pdfDef.SourceFileName}");
                            continue;
                        }

                        int pageNumber = ResolvePdfPageNumber(pdfRef, pdfDef);
                        string pngPath = ConvertPdfToPng(resolvedPdfPath, ed, pageNumber);
                        if (string.IsNullOrEmpty(pngPath) || !File.Exists(pngPath))
                        {
                            ed.WriteMessage($"\nFailed to convert PDF '{pdfDef.SourceFileName}' to PNG. Skipping.");
                            continue;
                        }

                        var placement = BuildPlacementForPdf(pdfRef, pdfDef, pngPath, resolvedPdfPath, pageNumber, ed);
                        if (placement == null)
                        {
                            try { File.Delete(pngPath); } catch { }
                            continue;
                        }

                        placement.TargetBtrId = btrId;
                        placement.Source = PlacementSource.Pdf;
                        _pending.Enqueue(placement);
                        queued++;
                    }
                }

                tr.Commit();
            }

            return queued;
        }

        private static ImagePlacement BuildPlacementForPdf(
            UnderlayReference pdfRef,
            UnderlayDefinition pdfDef,
            string pngPath,
            string resolvedPdfPath,
            int pageNumber,
            Editor ed)
        {
            var transform = pdfRef.Transform;
            double rawMinX = double.NaN, rawMinY = double.NaN, rawMaxX = double.NaN, rawMaxY = double.NaN;
            double leftPercent = double.NaN, bottomPercent = double.NaN, rightPercent = double.NaN, topPercent = double.NaN;
            bool clipDerived = false;
            string FormatPerc(double value) => double.IsNaN(value) || double.IsInfinity(value) ? "NaN" : value.ToString("F4");

            Point3d origin = transform.CoordinateSystem3d.Origin;
            Vector3d uVec = transform.CoordinateSystem3d.Xaxis;
            Vector3d vVec = transform.CoordinateSystem3d.Yaxis;

            Point3d placementOrigin = origin;
            Vector3d placementU = uVec;
            Vector3d placementV = vVec;
            Point2d[] percentageClipBoundary = null;

            bool isEffectivelyClipped = false;
            Point2d[] boundaryPoints = null;

            if (pdfRef.IsClipped)
            {
                try
                {
                    object rawBoundary = pdfRef.GetClipBoundary();
                    if (rawBoundary is Point2dCollection collection)
                    {
                        boundaryPoints = new Point2d[collection.Count];
                        collection.CopyTo(boundaryPoints, 0);
                    }
                    else if (rawBoundary is Point2d[] array)
                    {
                        boundaryPoints = array;
                    }

                    if (boundaryPoints != null && boundaryPoints.Length >= 2)
                    {
                        Point2d? sheetSize = GetPdfPageSizeInches(resolvedPdfPath, pageNumber);
                        if (sheetSize.HasValue)
                        {
                            double minClipX = double.PositiveInfinity, minClipY = double.PositiveInfinity;
                            double maxClipX = double.NegativeInfinity, maxClipY = double.NegativeInfinity;
                            foreach (var pt in boundaryPoints)
                            {
                                if (pt.X < minClipX) minClipX = pt.X;
                                if (pt.Y < minClipY) minClipY = pt.Y;
                                if (pt.X > maxClipX) maxClipX = pt.X;
                                if (pt.Y > maxClipY) maxClipY = pt.Y;
                            }

                            double pageWidth = sheetSize.Value.X;
                            double pageHeight = sheetSize.Value.Y;
                            double tolerance = 1e-4;

                            bool isDefaultBoundary =
                                Math.Abs(minClipX) < tolerance &&
                                Math.Abs(minClipY) < tolerance &&
                                Math.Abs(maxClipX - pageWidth) < tolerance &&
                                Math.Abs(maxClipY - pageHeight) < tolerance;

                            if (!isDefaultBoundary)
                                isEffectivelyClipped = true;
                        }
                        else
                        {
                            isEffectivelyClipped = true;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Failed to evaluate PDF clip boundary to check for clipping status: {ex.Message}");
                    isEffectivelyClipped = true;
                }
            }

            if (isEffectivelyClipped && boundaryPoints != null)
            {
                try
                {
                    double minX = double.PositiveInfinity;
                    double minY = double.PositiveInfinity;
                    double maxX = double.NegativeInfinity;
                    double maxY = double.NegativeInfinity;

                    foreach (var pt in boundaryPoints)
                    {
                        if (pt.X < minX) minX = pt.X;
                        if (pt.Y < minY) minY = pt.Y;
                        if (pt.X > maxX) maxX = pt.X;
                        if (pt.Y > maxY) maxY = pt.Y;
                    }

                    if (minX <= maxX && minY <= maxY)
                    {
                        rawMinX = minX;
                        rawMinY = minY;
                        rawMaxX = maxX;
                        rawMaxY = maxY;

                        Point3d localBottomLeft = new Point3d(minX, minY, 0);
                        Point3d localBottomRight = new Point3d(maxX, minY, 0);
                        Point3d localTopLeft = new Point3d(minX, maxY, 0);

                        Point3d worldBottomLeft = localBottomLeft.TransformBy(transform);
                        Point3d worldBottomRight = localBottomRight.TransformBy(transform);
                        Point3d worldTopLeft = localTopLeft.TransformBy(transform);

                        placementOrigin = worldBottomLeft;
                        placementU = worldBottomRight - worldBottomLeft;
                        placementV = worldTopLeft - worldBottomLeft;

                        var clipResult = ComputeClipPercentages(
                            resolvedPdfPath, pageNumber,
                            minX, minY, maxX, maxY,
                            transform, origin, uVec, vVec, ed);

                        percentageClipBoundary = clipResult.clip;
                        leftPercent = clipResult.left;
                        bottomPercent = clipResult.bottom;
                        rightPercent = clipResult.right;
                        topPercent = clipResult.top;

                        clipDerived = true;

                        if (placementU.Length < 1e-6 || placementV.Length < 1e-6)
                        {
                            ed.WriteMessage("\nSkipping PDF underlay with zero-area clip boundary.");
                            return null;
                        }

                        if (percentageClipBoundary == null)
                        {
                            ed.WriteMessage("\nWarning: Could not derive crop percentages for PowerPoint. The image may not be cropped correctly.");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Failed to evaluate PDF clip boundary: {ex.Message}");
                }
            }
            else
            {
                try
                {
                    Point2d? sheetSize = GetPdfPageSizeInches(resolvedPdfPath, pageNumber);
                    if (sheetSize.HasValue)
                    {
                        double sheetWidth = sheetSize.Value.X;
                        double sheetHeight = sheetSize.Value.Y;

                        Point3d localBottomLeft = new Point3d(0, 0, 0);
                        Point3d localBottomRight = new Point3d(sheetWidth, 0, 0);
                        Point3d localTopLeft = new Point3d(0, sheetHeight, 0);

                        Point3d worldBottomLeft = localBottomLeft.TransformBy(transform);
                        Point3d worldBottomRight = localBottomRight.TransformBy(transform);
                        Point3d worldTopLeft = localTopLeft.TransformBy(transform);

                        placementOrigin = worldBottomLeft;
                        placementU = worldBottomRight - worldBottomLeft;
                        placementV = worldTopLeft - worldBottomLeft;
                    }
                    else
                    {
                        ed.WriteMessage($"\nWarning: Could not determine original sheet size for '{Path.GetFileName(resolvedPdfPath)}'. Placement may be incorrect.");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Failed to calculate placement for unclipped PDF: {ex.Message}");
                }
            }

            string fileLabel = string.IsNullOrEmpty(resolvedPdfPath) ? "<unknown>" : Path.GetFileName(resolvedPdfPath);
            string clipDebug = clipDerived
                ? $"clip rawX=[{FormatPerc(rawMinX)},{FormatPerc(rawMaxX)}] rawY=[{FormatPerc(rawMinY)},{FormatPerc(rawMaxY)}] perc=[L={FormatPerc(leftPercent)},B={FormatPerc(bottomPercent)},R={FormatPerc(rightPercent)},T={FormatPerc(topPercent)}]"
                : "clip none";
            // Debug line omitted intentionally.

            if (placementU.Length < 1e-8 || placementV.Length < 1e-8)
            {
                ed.WriteMessage("\nSkipping PDF underlay due to zero-sized placement vectors.");
                return null;
            }

            return new ImagePlacement
            {
                Path = pngPath,
                Pos = placementOrigin,
                U = placementU,
                V = placementV,
                OriginalEntityId = pdfRef.ObjectId,
                ClipBoundary = percentageClipBoundary,
                Source = PlacementSource.Pdf
            };
        }

        private static string ConvertPdfToPng(string pdfPath, Editor ed, int pageNumber = 1)
        {
            try
            {
                int safePageNumber = pageNumber > 0 ? pageNumber : 1;
                int pageIndex = safePageNumber - 1;
                string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "PdfToPng");
                Directory.CreateDirectory(tempDir);
                string pngPath = Path.Combine(
                    tempDir,
                    $"{Path.GetFileNameWithoutExtension(pdfPath)}_p{safePageNumber:D3}_{Guid.NewGuid().ToString("N").Substring(0, 8)}.png"
                );
                var settings = new MagickReadSettings { Density = new Density(300, 300) };
                ed.WriteMessage($"\n[INFO] Converting PDF page {safePageNumber} to PNG using Magick.NET at 300 DPI...");
                using (var images = new MagickImageCollection())
                {
                    string pageSpecifier = $"{pdfPath}[{pageIndex}]";
                    images.Read(pageSpecifier, settings);
                    if (images.Count == 0)
                    {
                        ed.WriteMessage($"\nCould not read page {safePageNumber} from PDF: {pdfPath}.");
                        return null;
                    }
                    using (var pageImage = (MagickImage)images[0].Clone())
                    {
                        pageImage.Format = MagickFormat.Png;
                        pageImage.BackgroundColor = MagickColors.White;
                        pageImage.Alpha(AlphaOption.Remove);
                        pageImage.Write(pngPath);
                    }
                }
                if (File.Exists(pngPath) && new FileInfo(pngPath).Length > 0)
                {
                    try
                    {
                        using (var image = new MagickImage(pngPath))
                        {
                            if (image.Width > 0 && image.Height > 0)
                            {
                                ed.WriteMessage($"\nSuccessfully created PNG for page {safePageNumber}: {image.Width}x{image.Height} pixels, {new FileInfo(pngPath).Length} bytes");
                                return pngPath;
                            }
                            ed.WriteMessage("\nConversion failed: PNG has invalid dimensions.");
                            File.Delete(pngPath);
                            return null;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nPNG validation failed: {ex.Message}");
                        File.Delete(pngPath);
                        return null;
                    }
                }
                ed.WriteMessage("\nConversion failed: Output PNG is empty or missing.");
                if (File.Exists(pngPath)) File.Delete(pngPath);
                return null;
            }
            catch (MagickException ex)
            {
                ed.WriteMessage($"\n[ERROR] Magick.NET failed: {ex.Message}");
                ed.WriteMessage("\n[HINT] Magick.NET requires Ghostscript (gswin64c.exe) in PATH for PDF input.");
                return null;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPDF to PNG conversion failed: {ex.Message}");
                return null;
            }
        }

        private static int ResolvePdfPageNumber(UnderlayReference pdfRef, UnderlayDefinition pdfDef)
        {
            int page = TryParsePdfPage(SafeGetItemName(pdfRef));
            if (page > 0) return page;

            page = TryParsePdfPage(pdfDef?.ItemName);
            if (page > 0) return page;

            return 1;
        }

        private static string SafeGetItemName(UnderlayReference pdfRef)
        {
            if (pdfRef == null) return null;
            try
            {
                var prop = pdfRef.GetType().GetProperty("ItemName");
                if (prop != null && prop.PropertyType == typeof(string))
                    return prop.GetValue(pdfRef) as string;
            }
            catch { }
            return null;
        }

        private static int TryParsePdfPage(string candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate)) return -1;
            candidate = candidate.Trim();

            if (int.TryParse(candidate, out int direct) && direct > 0)
                return direct;

            int value = 0;
            bool foundDigit = false;

            foreach (char c in candidate)
            {
                if (char.IsDigit(c))
                {
                    value = (value * 10) + (c - '0');
                    foundDigit = true;
                }
                else if (foundDigit)
                {
                    break;
                }
            }

            return (foundDigit && value > 0) ? value : -1;
        }

        private static (Point2d[] clip, double left, double bottom, double right, double top) ComputeClipPercentages(
            string resolvedPdfPath,
            int pageNumber,
            double minX,
            double minY,
            double maxX,
            double maxY,
            Matrix3d transform,
            Point3d origin,
            Vector3d uVec,
            Vector3d vVec,
            Editor ed)
        {
            Point2d[] clip = null;
            double leftPercent = double.NaN;
            double bottomPercent = double.NaN;
            double rightPercent = double.NaN;
            double topPercent = double.NaN;

            Point2d? sheetSize = GetPdfPageSizeInches(resolvedPdfPath, pageNumber);
            if (sheetSize.HasValue && sheetSize.Value.X > 1e-6 && sheetSize.Value.Y > 1e-6)
            {
                double sheetWidth = sheetSize.Value.X;
                double sheetHeight = sheetSize.Value.Y;

                leftPercent = Clamp01(minX / sheetWidth);
                bottomPercent = Clamp01(minY / sheetHeight);
                rightPercent = Clamp01(maxX / sheetWidth);
                topPercent = Clamp01(maxY / sheetHeight);

                clip = new[]
                {
                    new Point2d(leftPercent, bottomPercent),
                    new Point2d(rightPercent, topPercent)
                };
            }

            if (clip == null)
            {
                try
                {
                    Matrix3d inverse = transform.Inverse();
                    Point3d originLocal = origin.TransformBy(inverse);
                    Point3d xLocal = (origin + uVec).TransformBy(inverse);
                    Point3d yLocal = (origin + vVec).TransformBy(inverse);

                    double sheetWidthLocal = Math.Abs(xLocal.X - originLocal.X);
                    double sheetHeightLocal = Math.Abs(yLocal.Y - originLocal.Y);

                    if (sheetWidthLocal > 1e-6 && sheetHeightLocal > 1e-6)
                    {
                        leftPercent = Clamp01((minX - originLocal.X) / sheetWidthLocal);
                        bottomPercent = Clamp01((minY - originLocal.Y) / sheetHeightLocal);
                        rightPercent = Clamp01((maxX - originLocal.X) / sheetWidthLocal);
                        topPercent = Clamp01((maxY - originLocal.Y) / sheetHeightLocal);

                        clip = new[]
                        {
                            new Point2d(leftPercent, bottomPercent),
                            new Point2d(rightPercent, topPercent)
                        };
                    }
                }
                catch (System.Exception ex)
                {
                    ed?.WriteMessage($"\n[DEBUG] Failed to derive fallback clip percentages: {ex.Message}");
                }
            }

            if (clip == null)
                return (null, double.NaN, double.NaN, double.NaN, double.NaN);

            return (clip, leftPercent, bottomPercent, rightPercent, topPercent);
        }

        private static double Clamp01(double value)
        {
            if (double.IsNaN(value)) return 0.0;
            if (value < 0.0) return 0.0;
            if (value > 1.0) return 1.0;
            return value;
        }

        // Reuse from SimplerCommands
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
                catch { }
            }
        }

        private static void RestoreOriginalLayer(Database db, ObjectId originalClayer)
        {
            if (db != null && !originalClayer.IsNull && db.Clayer != originalClayer)
            {
                try
                {
                    using (var tr = db.TransactionManager.StartOpenCloseTransaction())
                    {
                        var ltr = (LayerTableRecord)tr.GetObject(originalClayer, OpenMode.ForRead);
                        if (ltr != null && !ltr.IsErased && !ltr.IsFrozen)
                        {
                            db.Clayer = originalClayer;
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    AutoCADApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nCould not restore original layer: {ex.Message}");
                }
            }
        }
    }
}