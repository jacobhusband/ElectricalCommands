using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Powerpoint = Microsoft.Office.Interop.PowerPoint;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("EMBEDFROMPDF", CommandFlags.Modal)]
        public static void EmbedFromPdf()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage("\n--- Starting EMBEDFROMPDF (Debug Mode) ---");

            _pending.Clear();
            _isEmbeddingProcessActive = false;
            ObjectId originalClayer = db.Clayer;
            Matrix3d originalUcs = ed.CurrentUserCoordinateSystem;

            try
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity; // Set UCS to World

                ImagePlacement placementToDebug = null;

                // --- Phase 1: Find the FIRST PDF and prepare it ---
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // Find first PDF
                    (ObjectId refId, ObjectId ownerId) firstPdf = default;
                    bool found = false;
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null) continue;

                        foreach (ObjectId entId in btr)
                        {
                            if (entId.ObjectClass.DxfName.Equals("PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                            {
                                firstPdf = (entId, btr.ObjectId);
                                found = true;
                                break;
                            }
                        }
                        if (found) break;
                    }

                    if (!found)
                    {
                        ed.WriteMessage("\nNo valid PDF underlays found to embed.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\nFound a PDF underlay to process for debugging.");

                    var pdfRef = tr.GetObject(firstPdf.refId, OpenMode.ForRead) as UnderlayReference;
                    if (pdfRef == null) { tr.Commit(); return; }

                    var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                    if (pdfDef == null) { tr.Commit(); return; }

                    // Get PDF page size in drawing units from the reference transform
                    Matrix3d pdfTransform = pdfRef.Transform;
                    Vector3d u_vec = pdfTransform.CoordinateSystem3d.Xaxis;
                    Vector3d v_vec = pdfTransform.CoordinateSystem3d.Yaxis;

                    double pdfWidthInUnits = u_vec.Length;
                    double pdfHeightInUnits = v_vec.Length;

                    ed.WriteMessage($"\n[DEBUG] PDF Size (drawing units): W={pdfWidthInUnits:F4}, H={pdfHeightInUnits:F4}");

                    string resolvedPdfPath = ResolveImagePath(db, pdfDef.SourceFileName);
                    if (string.IsNullOrEmpty(resolvedPdfPath) || !File.Exists(resolvedPdfPath))
                    {
                        ed.WriteMessage($"\nSkipping missing PDF file: {pdfDef.SourceFileName}");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\nProcessing {Path.GetFileName(resolvedPdfPath)}...");
                    string pngPath = ConvertPdfToPng(resolvedPdfPath, ed);
                    if (string.IsNullOrEmpty(pngPath))
                    {
                        ed.WriteMessage("\nFailed to convert PDF to PNG. Skipping.");
                        tr.Commit();
                        return;
                    }

                    // --- NEW: Get Clip Boundary and Convert to Percentages ---
                    Point2d[] percentageClipBoundary = null;

                    if (pdfRef.IsClipped)
                    {
                        ed.WriteMessage("\nPDF is clipped. Calculating percentage-based clip boundary...");

                        object rawClipBoundaryObject = pdfRef.GetClipBoundary();
                        Point2d[] wcsClipBoundary = null;

                        if (rawClipBoundaryObject is Point2dCollection rawClipBoundaryCollection)
                        {
                            wcsClipBoundary = new Point2d[rawClipBoundaryCollection.Count];
                            rawClipBoundaryCollection.CopyTo(wcsClipBoundary, 0);
                        }
                        else if (rawClipBoundaryObject is Point2d[] rawClipBoundaryArray)
                        {
                            wcsClipBoundary = rawClipBoundaryArray;
                        }

                        if (wcsClipBoundary != null && wcsClipBoundary.Length >= 2)
                        {
                            // Get min/max of WCS clip boundary (these are in drawing units)
                            double minX_wcs = wcsClipBoundary.Min(p => p.X);
                            double minY_wcs = wcsClipBoundary.Min(p => p.Y);
                            double maxX_wcs = wcsClipBoundary.Max(p => p.X);
                            double maxY_wcs = wcsClipBoundary.Max(p => p.Y);

                            ed.WriteMessage($"\n[DEBUG] WCS Clip (units): MinX={minX_wcs:F4}, MinY={minY_wcs:F4}, MaxX={maxX_wcs:F4}, MaxY={maxY_wcs:F4}");

                            // Get the actual PDF sheet size in inches using Spire.PDF
                            int pageNumber = 1;
                            if (int.TryParse(pdfDef.ItemName, out int pg) && pg > 0)
                            {
                                pageNumber = pg;
                            }

                            Point2d? pdfSheetSizeInches = GetPdfPageSizeInches(resolvedPdfPath, pageNumber);

                            if (!pdfSheetSizeInches.HasValue)
                            {
                                ed.WriteMessage("\nWarning: Could not determine PDF sheet size. Cannot calculate crop percentages.");
                            }
                            else
                            {
                                double sheetWidthInches = pdfSheetSizeInches.Value.X;
                                double sheetHeightInches = pdfSheetSizeInches.Value.Y;

                                ed.WriteMessage($"\n[DEBUG] PDF Sheet Size (inches): W={sheetWidthInches:F4}\", H={sheetHeightInches:F4}\"");

                                // Calculate the scale factor: drawing units per inch
                                double unitsPerInchX = pdfWidthInUnits / sheetWidthInches;
                                double unitsPerInchY = pdfHeightInUnits / sheetHeightInches;

                                ed.WriteMessage($"\n[DEBUG] Scale Factor: {unitsPerInchX:F4} units/inch (X), {unitsPerInchY:F4} units/inch (Y)");

                                // Get the PDF origin in WCS
                                Point3d pdfOrigin = pdfTransform.CoordinateSystem3d.Origin;
                                ed.WriteMessage($"\n[DEBUG] PDF Origin (WCS): ({pdfOrigin.X:F4}, {pdfOrigin.Y:F4})");

                                // Transform the clip boundary from WCS into the PDF's local coordinate system
                                Matrix3d inverseTransform = pdfTransform.Inverse();

                                double minLocalX = double.MaxValue;
                                double minLocalY = double.MaxValue;
                                double maxLocalX = double.MinValue;
                                double maxLocalY = double.MinValue;

                                foreach (var clipPt in wcsClipBoundary)
                                {
                                    Point3d localPt = new Point3d(clipPt.X, clipPt.Y, 0).TransformBy(inverseTransform);
                                    if (localPt.X < minLocalX) minLocalX = localPt.X;
                                    if (localPt.Y < minLocalY) minLocalY = localPt.Y;
                                    if (localPt.X > maxLocalX) maxLocalX = localPt.X;
                                    if (localPt.Y > maxLocalY) maxLocalY = localPt.Y;
                                }

                                ed.WriteMessage($"\n[DEBUG] Clip in PDF Local Space (units): MinU={minLocalX:F4}, MinV={minLocalY:F4}, MaxU={maxLocalX:F4}, MaxV={maxLocalY:F4}");

                                double clipMinX_inches = minLocalX / unitsPerInchX;
                                double clipMinY_inches = minLocalY / unitsPerInchY;
                                double clipMaxX_inches = maxLocalX / unitsPerInchX;
                                double clipMaxY_inches = maxLocalY / unitsPerInchY;

                                clipMinX_inches = Math.Max(0.0, Math.Min(sheetWidthInches, clipMinX_inches));
                                clipMaxX_inches = Math.Max(0.0, Math.Min(sheetWidthInches, clipMaxX_inches));
                                clipMinY_inches = Math.Max(0.0, Math.Min(sheetHeightInches, clipMinY_inches));
                                clipMaxY_inches = Math.Max(0.0, Math.Min(sheetHeightInches, clipMaxY_inches));

                                if (clipMaxX_inches < clipMinX_inches)
                                {
                                    double swap = clipMinX_inches;
                                    clipMinX_inches = clipMaxX_inches;
                                    clipMaxX_inches = swap;
                                }

                                if (clipMaxY_inches < clipMinY_inches)
                                {
                                    double swap = clipMinY_inches;
                                    clipMinY_inches = clipMaxY_inches;
                                    clipMaxY_inches = swap;
                                }

                                ed.WriteMessage($"\n[DEBUG] Clip in Inches (from origin): MinX={clipMinX_inches:F4}\", MinY={clipMinY_inches:F4}\", MaxX={clipMaxX_inches:F4}\", MaxY={clipMaxY_inches:F4}\"");

                                // Calculate percentages based on the sheet size
                                double leftPercent = minX_wcs / sheetWidthInches;
                                double bottomPercent = minY_wcs / sheetHeightInches;
                                double rightPercent = maxX_wcs / sheetWidthInches;
                                double topPercent = maxY_wcs / sheetHeightInches;

                                leftPercent = Math.Max(0.0, Math.Min(1.0, leftPercent));
                                bottomPercent = Math.Max(0.0, Math.Min(1.0, bottomPercent));
                                rightPercent = Math.Max(0.0, Math.Min(1.0, rightPercent));
                                topPercent = Math.Max(0.0, Math.Min(1.0, topPercent));

                                ed.WriteMessage($"\n[DEBUG] Clip Percentages: Left={leftPercent:P2}, Bottom={bottomPercent:P2}, Right={rightPercent:P2}, Top={topPercent:P2}");

                                // Store as Point2d array: [0] = min (left, bottom), [1] = max (right, top)
                                percentageClipBoundary = new Point2d[]
                                {
                            new Point2d(leftPercent, bottomPercent),
                            new Point2d(rightPercent, topPercent)
                                };
                            }
                        }
                        else
                        {
                            ed.WriteMessage("\nWarning: Could not retrieve valid clipping boundary.");
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nPDF is not clipped. No cropping will be applied.");
                    }

                    placementToDebug = new ImagePlacement
                    {
                        Path = pngPath,
                        Pos = pdfTransform.CoordinateSystem3d.Origin,
                        U = u_vec,
                        V = v_vec,
                        OriginalEntityId = pdfRef.ObjectId,
                        ClipBoundary = percentageClipBoundary // Store percentage-based boundary
                    };

                    tr.Commit();
                }

                // --- Phase 2: Show in PowerPoint and stop ---
                if (placementToDebug == null)
                {
                    ed.WriteMessage("\nCould not prepare PDF for debugging.");
                    return;
                }

                ed.WriteMessage($"\nOpening PowerPoint to show cropped image. The command will stop here.");
                ShowImageInPowerPointForDebugPercentage(placementToDebug, ed);
                ed.WriteMessage("\nPowerPoint should be open with the cropped image. Please inspect it.");
                ed.WriteMessage("\nCommand finished.");

                ed.WriteMessage("\nPasting cropped image into drawing...");
                _savedClayer = originalClayer;
                _pending.Enqueue(placementToDebug);
                _lastPastedOle = ObjectId.Null;
                _finalPastedOleForZoom = ObjectId.Null;
                AttachHandlers(db, doc);
                _isEmbeddingProcessActive = true;
                ProcessNextPaste(doc, ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[!] An error occurred during the PDF embedding debug process: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (db.Clayer != originalClayer) db.Clayer = originalClayer;
                }
                catch { { } }
                ed.CurrentUserCoordinateSystem = originalUcs;
            }
        }

        private static void ShowImageInPowerPointForDebugPercentage(ImagePlacement placement, Editor ed)
        {
            try
            {
                if (!EnsurePowerPoint(ed)) return;
                dynamic slide = _pptPresentationShared.Slides[1];
                Powerpoint.Shapes shapes = slide.Shapes;

                // Clear previous shapes
                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { { shapes[i].Delete(); } } catch { { } }
                }

                string path = placement.Path;
                Shape pic = shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10);

                // --- Percentage-Based Cropping Logic ---
                if (placement.ClipBoundary != null && placement.ClipBoundary.Length >= 2)
                {
                    ed.WriteMessage("\nApplying percentage-based clipping boundary in PowerPoint...");

                    pic.ScaleWidth(1.00f, MsoTriState.msoTrue);

                    // Get the size of the inserted picture in PowerPoint points
                    float picWidthInPoints = pic.Width;
                    float picHeightInPoints = pic.Height;

                    ed.WriteMessage($"\n[DEBUG] PPT Picture Dimensions (points): W={picWidthInPoints:F4}, H={picHeightInPoints:F4}");

                    // Extract percentages from ClipBoundary
                    // [0] = (leftPercent, bottomPercent), [1] = (rightPercent, topPercent)
                    double leftPercent = placement.ClipBoundary[0].X;
                    double bottomPercent = placement.ClipBoundary[0].Y;
                    double rightPercent = placement.ClipBoundary[1].X;
                    double topPercent = placement.ClipBoundary[1].Y;

                    ed.WriteMessage($"\n[DEBUG] Clip Percentages: Left={leftPercent:P2}, Bottom={bottomPercent:P2}, Right={rightPercent:P2}, Top={topPercent:P2}");

                    // Calculate crop amounts in points
                    // Crop from left = leftPercent * width
                    float cropLeft = (float)(leftPercent * picWidthInPoints);

                    // Crop from top = (1 - topPercent) * height
                    float cropTop = (float)((1.0 - topPercent) * picHeightInPoints);

                    // Crop from right = (1 - rightPercent) * width
                    float cropRight = (float)((1.0 - rightPercent) * picWidthInPoints);

                    // Crop from bottom = bottomPercent * height
                    float cropBottom = (float)(bottomPercent * picHeightInPoints);

                    // Apply the crop
                    var picFormat = pic.PictureFormat;
                    picFormat.CropLeft = cropLeft;
                    picFormat.CropTop = cropTop;
                    picFormat.CropRight = cropRight;
                    picFormat.CropBottom = cropBottom;

                    ed.WriteMessage($"\n[DEBUG] Crop values (pts): L={cropLeft:F2}, T={cropTop:F2}, R={cropRight:F2}, B={cropBottom:F2}");

                    pic.Copy();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to show image in PowerPoint for debug: {ex.Message}");
            }
        }
    }
}
