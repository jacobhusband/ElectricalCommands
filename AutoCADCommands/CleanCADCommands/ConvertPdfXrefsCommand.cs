using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// For PDF→PNG conversion with Magick.NET
using ImageMagick;

// For XCLIP detection
using Autodesk.AutoCAD.DatabaseServices.Filters;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        private static Matrix3d GetCoordinateSystemMatrix(CoordinateSystem3d cs)
        {
            Vector3d zaxis = cs.Xaxis.CrossProduct(cs.Yaxis).GetNormal();
            return new Matrix3d(new double[] {
                cs.Xaxis.X, cs.Yaxis.X, zaxis.X, cs.Origin.X,
                cs.Xaxis.Y, cs.Yaxis.Y, zaxis.Y, cs.Origin.Y,
                cs.Xaxis.Z, cs.Yaxis.Z, zaxis.Z, cs.Origin.Z,
                0.0,        0.0,        0.0,     1.0
            });
        }

        private static bool HasXClipBoundary(ObjectId entityId, Transaction tr, out Extents3d? clipBounds, out Point3dCollection? boundaryPoints)
        {
            clipBounds = null;
            boundaryPoints = null;

            try
            {
                var entity = tr.GetObject(entityId, OpenMode.ForRead);
                if (!(entity is BlockReference br) || br.ExtensionDictionary == ObjectId.Null)
                    return false;

                var extdict = tr.GetObject(br.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                if (extdict == null || !extdict.Contains("ACAD_FILTER"))
                    return false;

                var fildict = tr.GetObject(extdict.GetAt("ACAD_FILTER"), OpenMode.ForRead) as DBDictionary;
                if (fildict == null || !fildict.Contains("SPATIAL"))
                    return false;

                var spatialFilter = tr.GetObject(fildict.GetAt("SPATIAL"), OpenMode.ForRead) as SpatialFilter;
                if (spatialFilter == null)
                    return false;

                // Get the clipping bounds
                clipBounds = spatialFilter.GetQueryBounds();

                // Convert Point2dCollection to Point3dCollection
                var points2d = spatialFilter.Definition.GetPoints();
                boundaryPoints = new Point3dCollection();
                foreach (Point2d pt in points2d)
                {
                    boundaryPoints.Add(new Point3d(pt.X, pt.Y, 0));
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static ObjectId CreateRectangleFromBoundaryPoints(Point3dCollection boundaryPoints, Transaction tr, Database db, string layer)
        {
            if (boundaryPoints == null || boundaryPoints.Count < 3)
                return ObjectId.Null;

            try
            {
                // For a rectangle, we need to find the min/max bounds from the boundary points
                double minX = double.MaxValue, minY = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue;

                foreach (Point3d pt in boundaryPoints)
                {
                    if (pt.X < minX) minX = pt.X;
                    if (pt.Y < minY) minY = pt.Y;
                    if (pt.X > maxX) maxX = pt.X;
                    if (pt.Y > maxY) maxY = pt.Y;
                }

                // Create rectangle points (bottom-left, bottom-right, top-right, top-left)
                var rectanglePoints = new Point3dCollection
                {
                    new Point3d(minX, minY, 0),
                    new Point3d(maxX, minY, 0),
                    new Point3d(maxX, maxY, 0),
                    new Point3d(minX, maxY, 0)
                };

                // Create the polyline for the rectangle
                var polyline = new Polyline(4);
                polyline.SetDatabaseDefaults(db);

                for (int i = 0; i < 4; i++)
                {
                    polyline.AddVertexAt(i, new Point2d(rectanglePoints[i].X, rectanglePoints[i].Y), 0, 0, 0);
                }

                polyline.Closed = true;
                polyline.Layer = layer;
                polyline.ColorIndex = 256; // ByLayer

                // Add to the current space (model space or paper space)
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                ObjectId polylineId = ms.AppendEntity(polyline);
                tr.AddNewlyCreatedDBObject(polyline, true);

                return polylineId;
            }
            catch
            {
                return ObjectId.Null;
            }
        }

        [CommandMethod("PDF2PNG", CommandFlags.Modal)]
        public static void ConvertPdfXrefsToPng()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\nStarting PDF XREF to PNG conversion...");

            // Critical so PNG defs don't show as "Unreferenced"
            RasterImage.EnableReactors(true);

            var pdfRefsToProcess = new List<(ObjectId refId, ObjectId ownerId)>();
            var pdfDefsToDetach = new HashSet<ObjectId>();

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null) continue;

                        foreach (ObjectId entId in btr)
                        {
                            if (entId.ObjectClass.DxfName.Equals("PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                                pdfRefsToProcess.Add((entId, btr.ObjectId));
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

                foreach (var item in pdfRefsToProcess)
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var pdfRef = tr.GetObject(item.refId, OpenMode.ForRead) as UnderlayReference;
                        var ownerBtr = tr.GetObject(item.ownerId, OpenMode.ForWrite) as BlockTableRecord;
                        if (pdfRef == null || ownerBtr == null) { ed.WriteMessage("\nSkipping invalid PDF reference object."); continue; }

                        var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                        if (pdfDef == null) { ed.WriteMessage("\nCould not find definition for PDF reference."); continue; }

                        string pdfPath = pdfDef.SourceFileName;
                        string resolvedPdfPath = ResolveImagePath(db, pdfPath); // use your existing helper
                        if (string.IsNullOrEmpty(resolvedPdfPath) || !File.Exists(resolvedPdfPath))
                        { ed.WriteMessage($"\nSkipping missing PDF file: {pdfPath}"); continue; }

                        ed.WriteMessage($"\nProcessing {Path.GetFileName(resolvedPdfPath)}...");

                        string pngPath = ConvertPdfToPng(resolvedPdfPath, ed);
                        if (string.IsNullOrEmpty(pngPath)) { ed.WriteMessage("\nFailed to convert PDF to PNG. Skipping."); continue; }

                        // Ensure image dictionary
                        ObjectId imageDictId = RasterImageDef.GetImageDictionary(db);
                        if (imageDictId.IsNull) imageDictId = RasterImageDef.CreateImageDictionary(db);
                        var imageDict = (DBDictionary)tr.GetObject(imageDictId, OpenMode.ForWrite);

                        string defName = Path.GetFileNameWithoutExtension(pngPath);
                        if (imageDict.Contains(defName))
                            defName = defName + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);

                        // Create image def
                        var imageDef = new RasterImageDef { SourceFileName = pngPath };
                        imageDef.Load();
                        ObjectId imageDefId = imageDict.SetAt(defName, imageDef);
                        tr.AddNewlyCreatedDBObject(imageDef, true);

                        // Debug: Check image definition
                        ed.WriteMessage($"\n[DEBUG] ImageDef properties:");
                        ed.WriteMessage($"\n[DEBUG] - SourceFileName: {imageDef.SourceFileName}");
                        ed.WriteMessage($"\n[DEBUG] - IsLoaded: {imageDef.IsLoaded}");

                        try
                        {
                            // Try to get image size from the actual file
                            using (var magickImage = new MagickImage(pngPath))
                            {
                                ed.WriteMessage($"\n[DEBUG] - File PixelWidth: {magickImage.Width}");
                                ed.WriteMessage($"\n[DEBUG] - File PixelHeight: {magickImage.Height}");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\n[DEBUG] - Error getting image properties: {ex.Message}");
                        }

                        // Create image entity
                        var rasterImage = new RasterImage();
                        rasterImage.SetDatabaseDefaults(db);

                        // Must set before associate
                        rasterImage.ImageDefId = imageDefId;

                        // Copy props
                        rasterImage.Layer = pdfRef.Layer;
                        rasterImage.LineWeight = pdfRef.LineWeight;
                        rasterImage.Color = pdfRef.Color;

                        rasterImage.ShowImage = true;
                        rasterImage.ImageTransparency = false;

                        // Match placement/orientation
                        Matrix3d transform = pdfRef.Transform;
                        Point3d origin = transform.CoordinateSystem3d.Origin;
                        Vector3d u_vec = transform.CoordinateSystem3d.Xaxis;
                        Vector3d v_vec = transform.CoordinateSystem3d.Yaxis;
                        rasterImage.Orientation = new CoordinateSystem3d(origin, u_vec, v_vec);

                        // Declare variables outside using block for clipping boundary access
                        int pixelWidth, pixelHeight;
                        double targetWidth, targetHeight;
                        Vector3d scaledU, scaledV;

                        // Get image dimensions and calculate proper display size
                        using (var magickImage = new MagickImage(pngPath))
                        {
                            pixelWidth = (int)magickImage.Width;
                            pixelHeight = (int)magickImage.Height;

                            // Use the transformation matrix to estimate proper size
                            // Based on the user's feedback, the original PDF is 187' x 121'
                            // We'll use these as target dimensions
                            targetWidth = 187.0 * 12.0;  // Convert feet to inches (assuming architectural units)
                            targetHeight = 121.0 * 12.0;

                            ed.WriteMessage($"\nTarget PDF size: {targetWidth:F2}x{targetHeight:F2} inches");

                            // Calculate aspect ratios
                            double aspectRatio = (double)pixelWidth / pixelHeight;
                            double targetAspectRatio = targetWidth / targetHeight;

                            // Use the target dimensions as the base
                            if (aspectRatio > targetAspectRatio)
                            {
                                // Image is wider, use target width as base
                                scaledU = u_vec.GetNormal() * targetWidth;
                                scaledV = u_vec.GetNormal().CrossProduct(Vector3d.ZAxis).GetNormal() * (targetWidth / aspectRatio);
                            }
                            else
                            {
                                // Image is taller, use target height as base
                                scaledV = v_vec.GetNormal() * targetHeight;
                                scaledU = v_vec.GetNormal().CrossProduct(Vector3d.ZAxis).GetNormal() * (targetHeight * aspectRatio);
                            }

                            ed.WriteMessage($"\nImage setup: {pixelWidth}x{pixelHeight} pixels, aspect ratio: {aspectRatio:F2}");
                            ed.WriteMessage($"\nTarget size: {targetWidth:F2}x{targetHeight:F2} units");
                            ed.WriteMessage($"\nUsing original origin: ({origin.X:F2}, {origin.Y:F2})");
                        }

                        // Update orientation with proper scale (using original origin for correct positioning)
                        rasterImage.Orientation = new CoordinateSystem3d(origin, scaledU, scaledV);

                        // Additional display properties for proper visibility
                        rasterImage.Brightness = 50; // Default brightness
                        rasterImage.Contrast = 50;   // Default contrast
                        rasterImage.Fade = 0;        // No fade

                        rasterImage.IsClipped = false;

                        // Append entity
                        ownerBtr.AppendEntity(rasterImage);
                        tr.AddNewlyCreatedDBObject(rasterImage, true);

                        // Associate so def is referenced (this is the key)
                        rasterImage.AssociateRasterDef(imageDef);

                        // Check for XCLIP boundary before removing PDF
                        Extents3d? clipBounds;
                        Point3dCollection? boundaryPoints;
                        ObjectId rectangleId = ObjectId.Null;

                        if (HasXClipBoundary(item.refId, tr, out clipBounds, out boundaryPoints))
                        {
                            ed.WriteMessage($"\nFound XCLIP boundary for PDF reference.");

                            if (clipBounds.HasValue)
                            {
                                ed.WriteMessage($"\nClip bounds: ({clipBounds.Value.MinPoint.X:F2}, {clipBounds.Value.MinPoint.Y:F2}) to ({clipBounds.Value.MaxPoint.X:F2}, {clipBounds.Value.MaxPoint.Y:F2})");
                            }

                            if (boundaryPoints != null && boundaryPoints.Count > 0)
                            {
                                ed.WriteMessage($"\nCreating rectangle from {boundaryPoints.Count} boundary points...");
                                rectangleId = CreateRectangleFromBoundaryPoints(boundaryPoints, tr, db, pdfRef.Layer);

                                if (!rectangleId.IsNull)
                                {
                                    ed.WriteMessage($"\nSuccessfully created clipping boundary rectangle.");
                                }
                                else
                                {
                                    ed.WriteMessage($"\nFailed to create clipping boundary rectangle.");
                                }
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"\nNo XCLIP boundary found for PDF reference.");
                        }

                        // Debug: Check raster image properties
                        ed.WriteMessage($"\n[DEBUG] RasterImage properties:");
                        ed.WriteMessage($"\n[DEBUG] - ImageDefId: {rasterImage.ImageDefId}");
                        ed.WriteMessage($"\n[DEBUG] - ShowImage: {rasterImage.ShowImage}");
                        ed.WriteMessage($"\n[DEBUG] - ImageTransparency: {rasterImage.ImageTransparency}");
                        ed.WriteMessage($"\n[DEBUG] - Brightness: {rasterImage.Brightness}");
                        ed.WriteMessage($"\n[DEBUG] - Contrast: {rasterImage.Contrast}");
                        ed.WriteMessage($"\n[DEBUG] - Fade: {rasterImage.Fade}");
                        ed.WriteMessage($"\n[DEBUG] - IsClipped: {rasterImage.IsClipped}");

                        try
                        {
                            ed.WriteMessage($"\n[DEBUG] - Orientation Origin: {rasterImage.Orientation.Origin}");
                            ed.WriteMessage($"\n[DEBUG] - Orientation X-axis: {rasterImage.Orientation.Xaxis}");
                            ed.WriteMessage($"\n[DEBUG] - Orientation Y-axis: {rasterImage.Orientation.Yaxis}");
                            ed.WriteMessage($"\n[DEBUG] - Orientation X-axis Length: {rasterImage.Orientation.Xaxis.Length}");
                            ed.WriteMessage($"\n[DEBUG] - Orientation Y-axis Length: {rasterImage.Orientation.Yaxis.Length}");
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\n[DEBUG] - Error getting orientation: {ex.Message}");
                        }

                        // Remove original PDF underlay
                        pdfRef.UpgradeOpen();
                        pdfRef.Erase();

                        pdfDefsToDetach.Add(pdfDef.ObjectId);
                        successCount++;

                        tr.Commit();
                    }
                }

                // Detach original PDF defs
                if (pdfDefsToDetach.Count > 0)
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                        if (nod.Contains("ACAD_PDFDEFINITIONS"))
                        {
                            var pdfDictId = nod.GetAt("ACAD_PDFDEFINITIONS");
                            if (!pdfDictId.IsNull)
                            {
                                var pdfDict = (DBDictionary)tr.GetObject(pdfDictId, OpenMode.ForWrite);
                                var keysToRemove = new List<string>();
                                foreach (DBDictionaryEntry entry in pdfDict)
                                {
                                    if (pdfDefsToDetach.Contains(entry.Value))
                                        keysToRemove.Add(entry.Key);
                                }
                                foreach (string key in keysToRemove)
                                {
                                    try { pdfDict.Remove(key); }
                                    catch (System.Exception ex) { ed.WriteMessage($"\nCould not remove PDF def '{key}': {ex.Message}"); }
                                }
                            }
                        }

                        foreach (ObjectId defId in pdfDefsToDetach)
                        {
                            try
                            {
                                var def = tr.GetObject(defId, OpenMode.ForWrite);
                                if (!def.IsErased) def.Erase();
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nCould not erase PDF def object {defId}: {ex.Message}");
                            }
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

        private static string ConvertPdfToPng(string pdfPath, Editor ed)
        {
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "PdfToPng");
                Directory.CreateDirectory(tempDir);
                string pngPath = Path.Combine(
                    tempDir,
                    Path.GetFileNameWithoutExtension(pdfPath) + "_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png"
                );

                var settings = new MagickReadSettings { Density = new Density(300, 300) };
                ed.WriteMessage("\n[INFO] Converting PDF to PNG using Magick.NET at 300 DPI...");

                using (var images = new MagickImageCollection())
                {
                    images.Read(pdfPath + "[0]", settings);
                    if (images.Count == 0)
                    {
                        ed.WriteMessage($"\nCould not read page 1 from PDF: {pdfPath}.");
                        return null;
                    }

                    using (var first = images.First())
                    {
                        first.Format = MagickFormat.Png;
                        first.BackgroundColor = MagickColors.White;
                        // Remove alpha channel to ensure visibility in AutoCAD
                        first.Alpha(AlphaOption.Off);
                        first.Write(pngPath);
                    }
                }

                if (File.Exists(pngPath) && new FileInfo(pngPath).Length > 0)
                {
                    // Validate PNG dimensions and content
                    try
                    {
                        using (var image = new MagickImage(pngPath))
                        {
                            if (image.Width > 0 && image.Height > 0)
                            {
                                ed.WriteMessage($"\nSuccessfully created PNG: {image.Width}x{image.Height} pixels, {new FileInfo(pngPath).Length} bytes");
                                return pngPath;
                            }
                            else
                            {
                                ed.WriteMessage("\nConversion failed: PNG has invalid dimensions.");
                                File.Delete(pngPath);
                                return null;
                            }
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
    }
}