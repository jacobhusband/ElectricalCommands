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
        [CommandMethod("GETPDFCLIPPING", CommandFlags.Modal)]
        public static void GetPdfClippingBoundary()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect a clipped PDF underlay:",
                SingleOnly = true
            };
            var filter = new SelectionFilter(new[] {
                new TypedValue((int)DxfCode.Start, "PDFUNDERLAY")
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

                    if (ent is UnderlayReference ur)
                    {
                        if (!ur.IsClipped)
                        {
                            ed.WriteMessage("\nThe selected PDF is not clipped.");
                            return;
                        }

                        object boundaryObject = ur.GetClipBoundary();
                        Point2d[] boundary = null;

                        if (boundaryObject is Point2dCollection boundaryCollection)
                        {
                            boundary = new Point2d[boundaryCollection.Count];
                            boundaryCollection.CopyTo(boundary, 0);
                        }
                        else if (boundaryObject is Point2d[] boundaryArray)
                        {
                            boundary = boundaryArray;
                        }

                        if (boundary != null && boundary.Length >= 2)
                        {
                            Point2d minPoint = boundary[0];
                            Point2d maxPoint = boundary[1];
                            ed.WriteMessage($"\n[DEBUG] Raw WCS Clip Boundary (drawing units):  -> ({minPoint.X:F4}, {minPoint.Y:F4})  -> ({maxPoint.X:F4}, {maxPoint.Y:F4})");
                        }
                        else
                        {
                            ed.WriteMessage("\nCould not retrieve a valid clipping boundary (requires at least 2 points).");
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nSelected object is not a PDF underlay.");
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