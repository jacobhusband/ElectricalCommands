using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Linq;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("GETCLIPBOUNDARY", CommandFlags.Modal)]
        public static void GetClippingBoundaryInches()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect a clipped underlay (PDF, DWF, DGN, or Image):",
                SingleOnly = true
            };
            var filter = new SelectionFilter(new[] {
                new TypedValue((int)DxfCode.Start, "PDFUNDERLAY,DWFUNDERLAY,DGNUNDERLAY,IMAGE")
            });

            var sel = ed.GetSelection(opts, filter);
            if (sel.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n*Cancel*");
                return;
            }

            var id = sel.Value.GetObjectIds().FirstOrDefault();
            if (id.IsNull) return;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null)
                    {
                        ed.WriteMessage("\nSelected object is not a valid entity.");
                        return;
                    }

                    Point2d[] boundary = null;
                    string underlayType = "";
                    object boundaryObject = null;

                    if (ent is UnderlayReference ur)
                    {
                        if (!ur.IsClipped)
                        {
                            ed.WriteMessage("\nThe selected underlay is not clipped.");
                            return;
                        }
                        boundaryObject = ur.GetClipBoundary();
                        underlayType = ur.GetRXClass().DxfName.Replace("UNDERLAY", "");
                    }
                    else if (ent is RasterImage ri)
                    {
                        if (!ri.IsClipped)
                        {
                            ed.WriteMessage("\nThe selected image is not clipped.");
                            return;
                        }
                        boundaryObject = ri.GetClipBoundary();
                        underlayType = "Image";
                    }

                    if (boundaryObject is Point2dCollection boundaryCollection)
                    {
                        boundary = new Point2d[boundaryCollection.Count];
                        boundaryCollection.CopyTo(boundary, 0);
                    }
                    else if (boundaryObject is Point2d[] boundaryArray)
                    {
                        boundary = boundaryArray;
                    }

                    if (boundary == null || boundary.Length == 0)
                    {
                        ed.WriteMessage("\nCould not retrieve a valid clipping boundary.");
                        return;
                    }

                    ed.WriteMessage($"\nClipping boundary for {underlayType} in WCS coordinates (inches):");

                    // Get the current drawing units
                    var insUnits = db.Insunits;
                    double conversionFactor = 1.0; // Default to inches

                    // Find the conversion factor to inches
                    switch (insUnits)
                    {
                        case UnitsValue.Millimeters:
                            conversionFactor = 1.0 / 25.4;
                            break;
                        case UnitsValue.Centimeters:
                            conversionFactor = 1.0 / 2.54;
                            break;
                        case UnitsValue.Meters:
                            conversionFactor = 1.0 / 0.0254;
                            break;
                        case UnitsValue.Feet:
                            conversionFactor = 12.0;
                            break;
                            // Add other cases as needed
                    }


                    for (int i = 0; i < boundary.Length; i++)
                    {
                        Point2d pt = boundary[i];
                        ed.WriteMessage($"  Vertex {i + 1}: (X={pt.X * conversionFactor:F4}, Y={pt.Y * conversionFactor:F4})");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred: {ex.Message}");
            }
        }
    }
}
