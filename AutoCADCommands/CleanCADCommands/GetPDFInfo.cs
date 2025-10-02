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
                            ed.WriteMessage("\n[Info] Drawing units are 'Undefined'. Assuming 1 drawing unit = 1 inch.");
                        }
                        else
                        {
                            ed.WriteMessage($"\n[Info] Drawing units: {drawingUnits}. Using conversion factor to inches: {toInchesFactor:F6}");
                        }

                        // --- Position & Scale ---
                        ed.WriteMessage("\n--- Position & Scale ---");
                        Point3d position = pdfRef.Position;
                        Point3d positionInInches = new Point3d(
                            position.X * toInchesFactor,
                            position.Y * toInchesFactor,
                            position.Z * toInchesFactor
                        );
                        ed.WriteMessage($"  AutoCAD Position (X, Y, Z) (Inches): {positionInInches.X:F4}, {positionInInches.Y:F4}, {positionInInches.Z:F4}");

                        Scale3d scale = pdfRef.ScaleFactors;
                        ed.WriteMessage($"  Applied Scale Factor (X, Y): {scale.X:F4}, {scale.Y:F4}");

                        double rotation = pdfRef.Rotation * (180.0 / Math.PI);
                        ed.WriteMessage($"  Rotation (Degrees): {rotation:F4}");

                        // --- Dimensions ---
                        ed.WriteMessage("\n--- Dimensions ---");
                        try
                        {
                            var extents = pdfRef.GeometricExtents;
                            double scaledWidthDrawingUnits = extents.MaxPoint.X - extents.MinPoint.X;
                            double scaledHeightDrawingUnits = extents.MaxPoint.Y - extents.MinPoint.Y;

                            double scaledWidthInches = scaledWidthDrawingUnits * toInchesFactor;
                            double scaledHeightInches = scaledHeightDrawingUnits * toInchesFactor;

                            ed.WriteMessage($"\n  AutoCAD Scaled Width of Clipped Region (Inches): {scaledWidthInches:F4}");
                            ed.WriteMessage($"\n  AutoCAD Scaled Height of Clipped Region (Inches): {scaledHeightInches:F4}");

                            // Calculate unscaled dimensions by dividing by the scale factor
                            if (Math.Abs(scale.X) > 1e-9 && Math.Abs(scale.Y) > 1e-9)
                            {
                                double unscaledWidthInches = scaledWidthInches / scale.X;
                                double unscaledHeightInches = scaledHeightInches / scale.Y;
                                ed.WriteMessage($"\n  Original Unscaled Width of Clipped Region (Inches): {unscaledWidthInches:F4}");
                                ed.WriteMessage($"\n  Original Unscaled Height of Clipped Region (Inches): {unscaledHeightInches:F4}");
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception)
                        {
                            ed.WriteMessage("  Dimensions: Not available (object may be off-screen).");
                        }

                        // --- Clipping ---
                        ed.WriteMessage("\n--- Clipping ---");
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
                                ed.WriteMessage("\n  Clip Boundary Points (relative to original corner) (inches):");
                                ed.WriteMessage($"\n    Point 1: ({boundaryPoints[0].X:F4}, {boundaryPoints[0].Y:F4})");
                                ed.WriteMessage($"\n    Point 2: ({boundaryPoints[1].X:F4}, {boundaryPoints[1].Y:F4})");

                                // The min X and Y of the local boundary points ARE the offset in inches
                                double offsetX_inches = boundaryPoints.Min(p => p.X);
                                double offsetY_inches = boundaryPoints.Min(p => p.Y);

                                ed.WriteMessage("\n  Clip Offset from Original Corner Unscaled (Inches):");
                                ed.WriteMessage($"\n    Distance (X, Y): ({offsetX_inches:F4}, {offsetY_inches:F4})");

                                ed.WriteMessage("\n  AutoCAD Clip Offset from Original Corner Scaled (Inches):");
                                ed.WriteMessage($"\n    Distance (X, Y): ({(offsetX_inches * scale.X):F4}, {(offsetY_inches * scale.Y):F4})");
                            }
                        }

                        // --- Miscellaneous Information ---
                        ed.WriteMessage("\n--- Miscellaneous ---");
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