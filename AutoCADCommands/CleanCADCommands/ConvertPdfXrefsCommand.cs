// filepath: AutoCADCommands/CleanCADCommands/ConvertPdfXrefsCommand.cs
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
                        rasterImage.ImageTransparency = true;

                        // Match placement/orientation
                        Matrix3d transform = pdfRef.Transform;
                        Point3d origin = transform.CoordinateSystem3d.Origin;
                        Vector3d u_vec = transform.CoordinateSystem3d.Xaxis;
                        Vector3d v_vec = transform.CoordinateSystem3d.Yaxis;
                        rasterImage.Orientation = new CoordinateSystem3d(origin, u_vec, v_vec);

                        // Copy clip
                        rasterImage.IsClipped = pdfRef.IsClipped;
                        if (pdfRef.IsClipped)
                        {
                            var boundaryObject = pdfRef.GetClipBoundary();
                            Point2d[] pdfBoundary2d = boundaryObject as Point2d[];
                            if (pdfBoundary2d != null && pdfBoundary2d.Length > 1)
                            {
                                Point3dCollection wcsBoundary3d = new Point3dCollection();
                                Matrix3d pdfTransform = pdfRef.Transform;
                                foreach (Point2d p2d in pdfBoundary2d)
                                    wcsBoundary3d.Add(new Point3d(p2d.X, p2d.Y, 0.0).TransformBy(pdfTransform));

                                Matrix3d imageTransform = GetCoordinateSystemMatrix(rasterImage.Orientation).Inverse();
                                Point2dCollection imageBoundary2d = new Point2dCollection();
                                foreach (Point3d p3d in wcsBoundary3d)
                                {
                                    Point3d transformed3d = p3d.TransformBy(imageTransform);
                                    imageBoundary2d.Add(new Point2d(transformed3d.X, transformed3d.Y));
                                }

                                if (imageBoundary2d.Count > 0 &&
                                    imageBoundary2d[0].GetDistanceTo(imageBoundary2d[imageBoundary2d.Count - 1]) > 1e-6)
                                {
                                    imageBoundary2d.Add(imageBoundary2d[0]);
                                }

                                if (imageBoundary2d.Count > 2)
                                    rasterImage.SetClipBoundary(ClipBoundaryType.Poly, imageBoundary2d);
                            }
                        }

                        // Append entity
                        ownerBtr.AppendEntity(rasterImage);
                        tr.AddNewlyCreatedDBObject(rasterImage, true);

                        // Associate so def is referenced (this is the key)
                        rasterImage.AssociateRasterDef(imageDef);

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
                        first.BackgroundColor = MagickColors.None;
                        first.Alpha(AlphaOption.Set);
                        first.Write(pngPath);
                    }
                }

                if (File.Exists(pngPath) && new FileInfo(pngPath).Length > 0) return pngPath;

                ed.WriteMessage("\nConversion failed: Output PNG is empty or missing.");
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
