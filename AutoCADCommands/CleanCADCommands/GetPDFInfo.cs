using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Linq;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;


namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("GETPDFINFO", CommandFlags.Modal)]
        public static void GetPdfInfo()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect a PDF Underlay to get its information: ");
            peo.SetRejectMessage("\nInvalid selection. Please select a PDF Underlay.");
            peo.AddAllowedClass(typeof(PdfReference), true);

            var per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nOperation cancelled.");
                return;
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var selectedId = per.ObjectId;
                    var pdfRef = tr.GetObject(selectedId, OpenMode.ForRead) as PdfReference;

                    if (pdfRef != null)
                    {
                        var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;

                        ed.WriteMessage("\n--- PDF Underlay Information ---");

                        // --- Unit Conversion Setup ---
                        UnitsValue drawingUnits = db.Insunits;
                        double toInchesFactor = GetConversionFactorToInches(drawingUnits);

                        if (drawingUnits == UnitsValue.Undefined)
                        {
                            ed.WriteMessage("\n[Warning] Drawing units are 'Undefined'. Assuming 1 drawing unit = 1 inch.");
                        }
                        else
                        {
                            ed.WriteMessage($"\nDrawing units: {drawingUnits}. Using conversion factor to inches: {toInchesFactor:F6}");
                        }

                        // --- Geometry Information (Now in Inches) ---
                        ed.WriteMessage("\nGeometry:");
                        Point3d position = pdfRef.Position;
                        Point3d positionInInches = new Point3d(
                            position.X * toInchesFactor,
                            position.Y * toInchesFactor,
                            position.Z * toInchesFactor
                        );
                        ed.WriteMessage($"  Position (X, Y, Z) (Inches): {positionInInches.X:F4}, {positionInInches.Y:F4}, {positionInInches.Z:F4}");

                        Scale3d scale = pdfRef.ScaleFactors;
                        ed.WriteMessage($"  Scale (X, Y, Z): {scale.X:F4}, {scale.Y:F4}, {scale.Z:F4}");

                        double rotation = pdfRef.Rotation * (180.0 / Math.PI); // Convert radians to degrees
                        ed.WriteMessage($"  Rotation (Degrees): {rotation:F4}");

                        try
                        {
                            var extents = pdfRef.GeometricExtents;
                            double widthInDrawingUnits = extents.MaxPoint.X - extents.MinPoint.X;
                            double heightInDrawingUnits = extents.MaxPoint.Y - extents.MinPoint.Y;
                            double widthInches = widthInDrawingUnits * toInchesFactor;
                            double heightInches = heightInDrawingUnits * toInchesFactor;
                            ed.WriteMessage($"  Width (Inches): {widthInches:F4}");
                            ed.WriteMessage($"  Height (Inches): {heightInches:F4}");
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception)
                        {
                            ed.WriteMessage("  Width/Height: Not available (object may be off-screen).");
                        }

                        // --- Clipping Information ---
                        ed.WriteMessage($"  Is Clipped: {pdfRef.IsClipped}");
                        if (pdfRef.IsClipped)
                        {
                            Point2d[] boundaryPoints = null;
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
                                ed.WriteMessage("  Clip Boundary Points (Local PDF Coords):");
                                ed.WriteMessage($"    Point 1: ({boundaryPoints[0].X:F4}, {boundaryPoints[0].Y:F4})");
                                ed.WriteMessage($"    Point 2: ({boundaryPoints[1].X:F4}, {boundaryPoints[1].Y:F4})");

                                // Find the min X and min Y from the boundary points to define the bottom-left corner of the clip area
                                double minX_local = boundaryPoints.Min(p => p.X);
                                double minY_local = boundaryPoints.Min(p => p.Y);

                                // Create a 3D point representing the local clip origin
                                Point3d localClipOrigin = new Point3d(minX_local, minY_local, 0);

                                // Use the PDF's transform matrix to find where this local point ends up in the drawing's world coordinates
                                Matrix3d transform = pdfRef.Transform;
                                Point3d worldClipOrigin = localClipOrigin.TransformBy(transform);

                                // The distance vector is the difference between the clipped corner's world position and the original corner's world position
                                Vector3d offsetVectorDrawingUnits = worldClipOrigin - pdfRef.Position;

                                // Convert the final vector to inches
                                Vector3d offsetVectorInches = offsetVectorDrawingUnits * toInchesFactor;

                                ed.WriteMessage("  Offset from Original Corner to Clipped Corner:");
                                ed.WriteMessage($"    Distance (X, Y, Z) (Inches): ({offsetVectorInches.X:F4}, {offsetVectorInches.Y:F4}, {offsetVectorInches.Z:F4})");
                            }
                        }

                        // --- Miscellaneous Information ---
                        ed.WriteMessage("\nMisc:");
                        ed.WriteMessage($"  Layer: {pdfRef.Layer}");
                        ed.WriteMessage($"  Object ID: {pdfRef.ObjectId}");
                        if (pdfDef != null)
                        {
                            ed.WriteMessage($"  PDF Source File: {pdfDef.SourceFileName}");
                            ed.WriteMessage($"  Active File Path: {ResolveImagePath(db, pdfDef.SourceFileName)}");
                            ed.WriteMessage($"  Page Number: {ResolvePdfPageNumber(pdfRef, pdfDef)}");
                        }
                        ed.WriteMessage("---------------------------------");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred: {ex.Message}");
                }
                finally
                {
                    tr.Commit();
                }
            }
        }

        private static double GetConversionFactorToInches(UnitsValue drawingUnits)
        {
            switch (drawingUnits)
            {
                case UnitsValue.Inches:
                case UnitsValue.USSurveyInch:
                    return 1.0;
                case UnitsValue.Feet:
                case UnitsValue.USSurveyFeet:
                    return 12.0;
                case UnitsValue.Yards:
                case UnitsValue.USSurveyYard:
                    return 36.0;
                case UnitsValue.Miles:
                case UnitsValue.USSurveyMile:
                    return 5280.0 * 12.0;
                case UnitsValue.Millimeters:
                    return 1.0 / 25.4;
                case UnitsValue.Centimeters:
                    return 1.0 / 2.54;
                case UnitsValue.Meters:
                    return 1000.0 / 25.4;
                case UnitsValue.Kilometers:
                    return (1000.0 * 1000.0) / 25.4;
                case UnitsValue.Undefined:
                default:
                    return 1.0;
            }
        }
    }
}