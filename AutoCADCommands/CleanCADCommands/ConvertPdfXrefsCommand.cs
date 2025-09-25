using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// For the placeholder PDF-to-PNG conversion, these are needed.
// NOTE: These are still needed as System.Drawing.Bitmap/Graphics are used in other methods (e.g. SimplerCommands.PreflightRasterForPpt).
using System.Drawing;
using System.Drawing.Imaging;

// *** NEW USING DIRECTIVE FOR GHOSTSCRIPT.NET ***
using Ghostscript.NET;
using Ghostscript.NET.Rasterizer;

// REMOVE: using ImageMagick; // Removed ImageMagick dependency

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        // New helper to replace the missing CoordinateSystem3d.ToMatrix3d() extension method.
        private static Matrix3d GetCoordinateSystemMatrix(CoordinateSystem3d cs)
        {
            // Build a Matrix3d from the CoordinateSystem3d's components.
            // The matrix is composed of the Xaxis, Yaxis, and Zaxis in the columns 
            // of the rotation/scale part, and the Origin in the translation part.
            Vector3d zaxis = cs.Xaxis.CrossProduct(cs.Yaxis).GetNormal();

            return new Matrix3d(
                new double[] {
                    cs.Xaxis.X, cs.Yaxis.X, zaxis.X, cs.Origin.X,
                    cs.Xaxis.Y, cs.Yaxis.Y, zaxis.Y, cs.Origin.Y,
                    cs.Xaxis.Z, cs.Yaxis.Z, zaxis.Z, cs.Origin.Z,
                    0.0, 0.0, 0.0, 1.0
                });
        }


        /// <summary>
        /// Finds all PDF Underlay XREFs, converts them to PNG images,
        /// inserts the PNGs to match the original PDFs, and removes the PDFs.
        /// </summary>
        [CommandMethod("PDF2PNG", CommandFlags.Modal)]
        public static void ConvertPdfXrefsToPng()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\nStarting PDF XREF to PNG conversion...");

            var pdfRefsToProcess = new List<(ObjectId refId, ObjectId ownerId)>();
            var pdfDefsToDetach = new HashSet<ObjectId>();

            try
            {
                // --- Phase 1: Find all PDF references in the drawing ---
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // Iterate through Model Space and all Paper Space layouts
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null) continue;

                        foreach (ObjectId entId in btr)
                        {
                            if (entId.ObjectClass.DxfName.Equals("PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                            {
                                pdfRefsToProcess.Add((entId, btr.ObjectId));
                            }
                        }
                    }
                    tr.Commit();
                }

                if (pdfRefsToProcess.Count == 0)
                {
                    ed.WriteMessage("\nNo PDF XREFs found in the drawing.");
                    return;
                }

                ed.WriteMessage($"\nFound {pdfRefsToProcess.Count} PDF XREF(s) to process.");
                int successCount = 0;

                // --- Phase 2: Process each PDF reference ---
                foreach (var item in pdfRefsToProcess)
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var pdfRef = tr.GetObject(item.refId, OpenMode.ForRead) as UnderlayReference;
                        var ownerBtr = tr.GetObject(item.ownerId, OpenMode.ForWrite) as BlockTableRecord;

                        if (pdfRef == null || ownerBtr == null)
                        {
                            ed.WriteMessage($"\nSkipping invalid PDF reference object.");
                            continue;
                        }

                        var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                        if (pdfDef == null)
                        {
                            ed.WriteMessage("\nCould not find definition for PDF reference.");
                            continue;
                        }

                        // 1. Get PDF Filepath
                        string pdfPath = pdfDef.SourceFileName;
                        string resolvedPdfPath = ResolveImagePath(db, pdfPath);

                        if (string.IsNullOrEmpty(resolvedPdfPath) || !File.Exists(resolvedPdfPath))
                        {
                            ed.WriteMessage($"\nSkipping missing PDF file: {pdfPath}");
                            continue;
                        }

                        ed.WriteMessage($"\nProcessing {Path.GetFileName(resolvedPdfPath)}...");

                        // 2. Convert PDF to PNG image
                        // *** UPDATED CALL TO NEW GHOSTSCRIPT-BASED CONVERTER ***
                        string pngPath = ConvertPdfToPng(resolvedPdfPath, ed);
                        if (string.IsNullOrEmpty(pngPath))
                        {
                            ed.WriteMessage($"\nFailed to convert PDF to PNG. Skipping.");
                            continue;
                        }
                        // ******************************************************

                        // 3. Insert the new PNG
                        ObjectId imageDefId;
                        var imageDictId = RasterImageDef.GetImageDictionary(db);
                        if (imageDictId.IsNull)
                        {
                            imageDictId = RasterImageDef.CreateImageDictionary(db);
                        }

                        var imageDict = (DBDictionary)tr.GetObject(imageDictId, OpenMode.ForWrite);
                        string defName = Path.GetFileNameWithoutExtension(pngPath);
                        if (imageDict.Contains(defName))
                        {
                            defName = defName + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                        }

                        using (var imageDef = new RasterImageDef())
                        {
                            imageDef.SourceFileName = pngPath;
                            imageDef.Load();
                            imageDefId = imageDict.SetAt(defName, imageDef);
                            tr.AddNewlyCreatedDBObject(imageDef, true);
                        }

                        // 4. Create the RasterImage entity and match the PDF's geometry
                        using (var rasterImage = new RasterImage())
                        {
                            rasterImage.ImageDefId = imageDefId;
                            rasterImage.SetDatabaseDefaults(db);
                            rasterImage.Layer = pdfRef.Layer;
                            rasterImage.LineWeight = pdfRef.LineWeight;
                            rasterImage.Color = pdfRef.Color;

                            // Replicate the PDF's transformation (Position, Scale, Rotation)
                            Matrix3d transform = pdfRef.Transform;
                            Point3d origin = transform.CoordinateSystem3d.Origin;
                            Vector3d u_vec = transform.CoordinateSystem3d.Xaxis;
                            Vector3d v_vec = transform.CoordinateSystem3d.Yaxis;

                            rasterImage.Orientation = new CoordinateSystem3d(origin, u_vec, v_vec);

                            // 5. Crop the PNG to match the PDF clip boundary
                            rasterImage.IsClipped = pdfRef.IsClipped;
                            if (pdfRef.IsClipped)
                            {
                                // GetClipBoundary() returns Point2d[] in some environments.
                                var boundaryObject = pdfRef.GetClipBoundary();
                                Point2d[] pdfBoundary2d = boundaryObject as Point2d[];

                                // Check if the boundary is a valid array and has enough points.
                                if (pdfBoundary2d != null && pdfBoundary2d.Length > 1)
                                {
                                    // 1. Transform from PDF Object Space to WCS 
                                    Point3dCollection wcsBoundary3d = new Point3dCollection();
                                    Matrix3d pdfTransform = pdfRef.Transform;

                                    foreach (Point2d p2d in pdfBoundary2d)
                                    {
                                        // Assume Z=0 on PDF object's local plane.
                                        Point3d p3d = new Point3d(p2d.X, p2d.Y, 0.0);
                                        // Transform the local point into WCS.
                                        wcsBoundary3d.Add(p3d.TransformBy(pdfTransform));
                                    }

                                    // 2. Transform from WCS to the new RasterImage's Object Space
                                    // The transformation to Image Object Space is the inverse of the image's final orientation transform.
                                    Matrix3d imageTransform = GetCoordinateSystemMatrix(rasterImage.Orientation).Inverse();

                                    Point2dCollection imageBoundary2d = new Point2dCollection();
                                    foreach (Point3d p3d in wcsBoundary3d)
                                    {
                                        Point3d transformed3d = p3d.TransformBy(imageTransform);
                                        // The RasterImage SetClipBoundary expects 2D points on the image's local plane.
                                        imageBoundary2d.Add(new Point2d(transformed3d.X, transformed3d.Y));
                                    }

                                    // *** Final robustness check: The method requires the last point to be identical to the first point. ***
                                    if (imageBoundary2d.Count > 0 && imageBoundary2d[0].GetDistanceTo(imageBoundary2d[imageBoundary2d.Count - 1]) > 1e-6)
                                    {
                                        imageBoundary2d.Add(imageBoundary2d[0]);
                                    }

                                    if (imageBoundary2d.Count > 2)
                                    {
                                        // 3. Set the clip boundary
                                        rasterImage.SetClipBoundary(ClipBoundaryType.Poly, imageBoundary2d);
                                    }
                                }
                            }

                            ownerBtr.AppendEntity(rasterImage);
                            tr.AddNewlyCreatedDBObject(rasterImage, true);
                        }

                        // 6. Remove the original PDF reference
                        pdfRef.UpgradeOpen();
                        pdfRef.Erase();

                        pdfDefsToDetach.Add(pdfDef.ObjectId);
                        successCount++;

                        tr.Commit();
                    }
                }

                // --- Phase 3: Detach the original PDF definitions ---
                if (pdfDefsToDetach.Count > 0)
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                        var pdfDictId = nod.GetAt("ACAD_PDFDEFINITIONS");
                        if (!pdfDictId.IsNull)
                        {
                            var pdfDict = (DBDictionary)tr.GetObject(pdfDictId, OpenMode.ForWrite);
                            var keysToRemove = new List<string>();
                            foreach (DBDictionaryEntry entry in pdfDict)
                            {
                                if (pdfDefsToDetach.Contains(entry.Value))
                                {
                                    keysToRemove.Add(entry.Key);
                                }
                            }

                            foreach (string key in keysToRemove)
                            {
                                try { pdfDict.Remove(key); }
                                catch (System.Exception ex) { ed.WriteMessage($"\nCould not remove PDF def '{key}': {ex.Message}"); }
                            }
                        }

                        foreach (ObjectId defId in pdfDefsToDetach)
                        {
                            try
                            {
                                var def = tr.GetObject(defId, OpenMode.ForWrite);
                                def.Erase();
                            }
                            catch (System.Exception ex) { ed.WriteMessage($"\nCould not erase PDF def object {defId}: {ex.Message}"); }
                        }
                        tr.Commit();
                        ed.WriteMessage($"\nDetached {pdfDefsToDetach.Count} PDF definition(s).");
                    }
                }

                ed.WriteMessage($"\nSuccessfully converted {successCount} of {pdfRefsToProcess.Count} PDF XREF(s) to PNG images.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Converts the first page of a PDF to a PNG image file using Ghostscript.NET.
        /// </summary>
        private static string ConvertPdfToPng(string pdfPath, Editor ed)
        {
            try
            {
                // 1. Setup temporary path for the output PNG
                string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "PdfToPng");
                Directory.CreateDirectory(tempDir);
                string pngPath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(pdfPath) + "_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");

                // 2. Find the Ghostscript Native DLL
                GhostscriptVersionInfo gv = GhostscriptVersionInfo.GetLastInstalledVersion();
                if (gv == null)
                {
                    ed.WriteMessage("\n[ERROR] Ghostscript native library not found. Please ensure the native Ghostscript application is installed (e.g., from ghostscript.com).");
                    return null;
                }

                ed.WriteMessage($"\n[INFO] Converting PDF to PNG using Ghostscript.NET at 300 DPI...");

                // 3. Use GhostscriptRasterizer to perform the conversion
                using (var rasterizer = new GhostscriptRasterizer())
                {
                    int dpi = 300;
                    int pageNumber = 1; // Rasterize the first page

                    // Open the PDF using the resolved Ghostscript version info
                    rasterizer.Open(pdfPath, gv, false);

                    if (rasterizer.PageCount < pageNumber)
                    {
                        ed.WriteMessage($"\nCould not read page {pageNumber} from PDF: {pdfPath}. File seems empty.");
                        return null;
                    }

                    // Rasterize the first page to a System.Drawing.Image object
                    using (var pageImage = rasterizer.GetPage(dpi, pageNumber))
                    {
                        // Save the image as PNG
                        pageImage.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }

                // Return the path if the file exists and is valid
                if (File.Exists(pngPath) && new FileInfo(pngPath).Length > 0)
                {
                    return pngPath;
                }

                ed.WriteMessage($"\nConversion failed: Output PNG file is empty or missing.");
                return null;
            }
            catch (GhostscriptLibraryNotInstalledException)
            {
                ed.WriteMessage("\n[ERROR] Ghostscript native library not found. Ensure 32-bit or 64-bit Ghostscript is installed.");
                return null;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPDF to PNG conversion failed: {ex.Message}");
                return null;
            }
        }
    }
}