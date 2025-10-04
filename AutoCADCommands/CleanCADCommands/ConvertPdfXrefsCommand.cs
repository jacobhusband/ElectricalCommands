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

        // --- THIS IS THE FIX for the warning ---
        private static bool HasXClipBoundary(ObjectId entityId, Transaction tr, out Extents3d clipBounds, out Point3dCollection boundaryPoints)
        {
            clipBounds = new Extents3d(); // Initialize to default
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

        private static ObjectId CreateRedRectangleFromBoundaryPoints(Point3dCollection boundaryPoints, Transaction tr, Database db, ObjectId ownerBtrId)
        {
            if (boundaryPoints == null || boundaryPoints.Count < 2)
            {
                return ObjectId.Null;
            }

            try
            {
                // --- NEW: Ensure a dedicated "clipping" layer exists and is visible ---
                const string clippingLayerName = "clipping";
                ObjectId clippingLayerId = ObjectId.Null;
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (!lt.Has(clippingLayerName))
                {
                    // Layer doesn't exist, create it
                    lt.UpgradeOpen();
                    var newLayer = new LayerTableRecord
                    {
                        Name = clippingLayerName,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1), // Red
                        IsOff = false,
                        IsFrozen = false,
                        IsLocked = false,
                        LineWeight = LineWeight.LineWeight050 // Make it visible
                    };
                    clippingLayerId = lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }
                else
                {
                    // Layer exists, ensure it's visible and unlocked
                    clippingLayerId = lt[clippingLayerName];
                    var existingLayer = (LayerTableRecord)tr.GetObject(clippingLayerId, OpenMode.ForWrite);
                    if (existingLayer.IsOff) existingLayer.IsOff = false;
                    if (existingLayer.IsFrozen) existingLayer.IsFrozen = false;
                    if (existingLayer.IsLocked) existingLayer.IsLocked = false;
                }

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
                // --- MODIFIED: Assign to the new layer and set properties to ByLayer ---
                polyline.Layer = clippingLayerName;
                polyline.ColorIndex = 256; // ByLayer
                polyline.LineWeight = LineWeight.ByLayer;

                // Add to the correct space (model or paper) passed in via ownerBtrId
                var ownerSpace = (BlockTableRecord)tr.GetObject(ownerBtrId, OpenMode.ForWrite);
                ObjectId polylineId = ownerSpace.AppendEntity(polyline);
                tr.AddNewlyCreatedDBObject(polyline, true);

                System.Diagnostics.Debug.WriteLine($"DEBUG: Successfully created clipping rectangle on layer '{clippingLayerName}' with ObjectId: {polylineId}");

                return polylineId;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: Exception in CreateRedRectangleFromBoundaryPoints: {ex.Message}");
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

                        // Check for clip boundary and create red rectangle visualization
                        Point2d[] clipBoundary = pdfRef.GetClipBoundary();
                        ObjectId redRectangleId = ObjectId.Null;

                        if (clipBoundary != null && clipBoundary.Length > 0)
                        {
                            ed.WriteMessage($"\nFound clip boundary for PDF reference with {clipBoundary.Length} points.");

                            var boundaryPoints3d = new Point3dCollection();
                            foreach (Point2d pt in clipBoundary)
                            {
                                boundaryPoints3d.Add(new Point3d(pt.X, pt.Y, 0));
                            }

                            ed.WriteMessage($"\nCreating clipping boundary rectangle on layer 'clipping'...");
                            // --- MODIFIED: Simplified call to use the new layer logic ---
                            redRectangleId = CreateRedRectangleFromBoundaryPoints(boundaryPoints3d, tr, db, item.ownerId);

                            if (!redRectangleId.IsNull)
                            {
                                ed.WriteMessage($"\nSuccessfully created clipping boundary rectangle.");

                                // Zoom to the rectangle to make it visible
                                try
                                {
                                    var minPoint = new Point3d(clipBoundary[0].X, clipBoundary[0].Y, 0);
                                    var maxPoint = new Point3d(clipBoundary[1].X, clipBoundary[1].Y, 0);

                                    double padding = Math.Max(maxPoint.X - minPoint.X, maxPoint.Y - minPoint.Y) * 0.1;
                                    var zoomMin = new Point3d(minPoint.X - padding, minPoint.Y - padding, 0);
                                    var zoomMax = new Point3d(maxPoint.X + padding, maxPoint.Y + padding, 0);

                                    ed.WriteMessage($"\nDEBUG: Zooming to rectangle area: ({zoomMin.X:F2}, {zoomMin.Y:F2}) to ({zoomMax.X:F2}, {zoomMax.Y:F2})");
                                }
                                catch (System.Exception ex)
                                {
                                    ed.WriteMessage($"\nDEBUG: Could not zoom to rectangle: {ex.Message}");
                                }
                            }
                            else
                            {
                                ed.WriteMessage($"\nFailed to create clipping boundary rectangle.");
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"\nNo clip boundary found for PDF reference.");
                        }

                        string pdfPath = pdfDef.SourceFileName;
                        string resolvedPdfPath = ResolveImagePath(db, pdfPath);
                        if (string.IsNullOrEmpty(resolvedPdfPath) || !File.Exists(resolvedPdfPath))
                        { ed.WriteMessage($"\nSkipping missing PDF file: {pdfPath}"); continue; }

                        int pageNumber = ResolvePdfPageNumber(pdfRef, pdfDef);
                        ed.WriteMessage($"\nProcessing {Path.GetFileName(resolvedPdfPath)} (page {pageNumber})...");

                        string pngPath = ConvertPdfToPng(resolvedPdfPath, ed, pageNumber);
                        if (string.IsNullOrEmpty(pngPath)) { ed.WriteMessage("\nFailed to convert PDF to PNG. Skipping."); continue; }

                        ObjectId imageDictId = RasterImageDef.GetImageDictionary(db);
                        if (imageDictId.IsNull) imageDictId = RasterImageDef.CreateImageDictionary(db);
                        var imageDict = (DBDictionary)tr.GetObject(imageDictId, OpenMode.ForWrite);

                        string defName = Path.GetFileNameWithoutExtension(pngPath);
                        if (imageDict.Contains(defName))
                            defName = defName + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);

                        var imageDef = new RasterImageDef { SourceFileName = pngPath };
                        imageDef.Load();
                        ObjectId imageDefId = imageDict.SetAt(defName, imageDef);
                        tr.AddNewlyCreatedDBObject(imageDef, true);

                        ed.WriteMessage($"\n[DEBUG] ImageDef properties:");
                        ed.WriteMessage($"\n[DEBUG] - SourceFileName: {imageDef.SourceFileName}");
                        ed.WriteMessage($"\n[DEBUG] - IsLoaded: {imageDef.IsLoaded}");

                        try
                        {
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

                        var rasterImage = new RasterImage();
                        rasterImage.SetDatabaseDefaults(db);
                        rasterImage.ImageDefId = imageDefId;
                        rasterImage.Layer = pdfRef.Layer;
                        rasterImage.LineWeight = pdfRef.LineWeight;
                        rasterImage.Color = pdfRef.Color;
                        rasterImage.ShowImage = true;
                        rasterImage.ImageTransparency = false;

                        Matrix3d transform = pdfRef.Transform;
                        Point3d origin = transform.CoordinateSystem3d.Origin;
                        Vector3d u_vec = transform.CoordinateSystem3d.Xaxis;
                        Vector3d v_vec = transform.CoordinateSystem3d.Yaxis;
                        rasterImage.Orientation = new CoordinateSystem3d(origin, u_vec, v_vec);

                        int pixelWidth, pixelHeight;
                        double targetWidth, targetHeight;
                        Vector3d scaledU, scaledV;

                        using (var magickImage = new MagickImage(pngPath))
                        {
                            pixelWidth = (int)magickImage.Width;
                            pixelHeight = (int)magickImage.Height;

                            targetWidth = 187.0 * 12.0;
                            targetHeight = 121.0 * 12.0;

                            ed.WriteMessage($"\nTarget PDF size: {targetWidth:F2}x{targetHeight:F2} inches");

                            double aspectRatio = (double)pixelWidth / pixelHeight;
                            double targetAspectRatio = targetWidth / targetHeight;

                            if (aspectRatio > targetAspectRatio)
                            {
                                scaledU = u_vec.GetNormal() * targetWidth;
                                scaledV = u_vec.GetNormal().CrossProduct(Vector3d.ZAxis).GetNormal() * (targetWidth / aspectRatio);
                            }
                            else
                            {
                                scaledV = v_vec.GetNormal() * targetHeight;
                                scaledU = v_vec.GetNormal().CrossProduct(Vector3d.ZAxis).GetNormal() * (targetHeight * aspectRatio);
                            }

                            ed.WriteMessage($"\nImage setup: {pixelWidth}x{pixelHeight} pixels, aspect ratio: {aspectRatio:F2}");
                            ed.WriteMessage($"\nTarget size: {targetWidth:F2}x{targetHeight:F2} units");
                            ed.WriteMessage($"\nUsing original origin: ({origin.X:F2}, {origin.Y:F2})");
                        }

                        rasterImage.Orientation = new CoordinateSystem3d(origin, scaledU, scaledV);
                        rasterImage.Brightness = 50;
                        rasterImage.Contrast = 50;
                        rasterImage.Fade = 0;
                        rasterImage.IsClipped = false;

                        ownerBtr.AppendEntity(rasterImage);
                        tr.AddNewlyCreatedDBObject(rasterImage, true);
                        rasterImage.AssociateRasterDef(imageDef);

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

                        pdfRef.UpgradeOpen();
                        pdfRef.Erase();

                        pdfDefsToDetach.Add(pdfDef.ObjectId);
                        successCount++;

                        tr.Commit();
                    }
                }

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

        [CommandMethod("DrawPdfBoundaries")]
        public static void DrawPdfBoundaries()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            int boundaryCount = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (ObjectId entId in ms)
                {
                    if (entId.ObjectClass.DxfName.Equals("PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                    {
                        var pdfRef = tr.GetObject(entId, OpenMode.ForRead) as UnderlayReference;
                        if (pdfRef != null && pdfRef.IsClipped)
                        {
                            Point2d[] boundary = pdfRef.GetClipBoundary();
                            if (boundary != null && boundary.Length > 0)
                            {
                                var polyline = new Polyline(boundary.Length);
                                polyline.SetDatabaseDefaults(db);

                                for (int i = 0; i < boundary.Length; i++)
                                {
                                    Point3d pt = new Point3d(boundary[i].X, boundary[i].Y, 0);
                                    pt = pt.TransformBy(pdfRef.Transform);
                                    polyline.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
                                }

                                polyline.Closed = true;
                                polyline.ColorIndex = 1; // Red 

                                ms.AppendEntity(polyline);
                                tr.AddNewlyCreatedDBObject(polyline, true);
                                boundaryCount++;
                            }
                        }
                    }
                }
                tr.Commit();
            }
            ed.WriteMessage($"\nDrew {boundaryCount} PDF clip boundaries in red.");
        }


    }
}