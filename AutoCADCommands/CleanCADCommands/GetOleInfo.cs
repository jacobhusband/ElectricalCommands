using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("GETOLEINFO", CommandFlags.Modal)]
        public static void GetOleInfo()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            // CORRECTED: Use PromptEntityOptions for a single selection, which reliably accepts a message.
            var peo = new PromptEntityOptions("\nSelect an OLE object to get its information: ");
            peo.SetRejectMessage("\nInvalid selection. Please select an OLE object.");
            peo.AddAllowedClass(typeof(Ole2Frame), true);

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
                    var oleFrame = tr.GetObject(selectedId, OpenMode.ForRead) as Ole2Frame;

                    if (oleFrame != null)
                    {
                        ed.WriteMessage("\n--- OLE Object Information ---");

                        // --- Geometry Information ---
                        ed.WriteMessage("\nGeometry:");
                        Point3d position = oleFrame.Location;
                        ed.WriteMessage($"  Position (X, Y, Z): {position.X:F4}, {position.Y:F4}, {position.Z:F4}");

                        // Get the size of the OLE object from its geometric extents
                        try
                        {
                            var extents = oleFrame.GeometricExtents;
                            Point3d minPoint = extents.MinPoint;
                            Point3d maxPoint = extents.MaxPoint;
                            double width = maxPoint.X - minPoint.X;
                            double height = maxPoint.Y - minPoint.Y;
                            ed.WriteMessage($"  Width: {width:F4}");
                            ed.WriteMessage($"  Height: {height:F4}");
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception)
                        {
                            ed.WriteMessage("  Width: Not available (object may be off-screen or have no extents).");
                            ed.WriteMessage("  Height: Not available.");
                        }

                        // --- Miscellaneous Information ---
                        ed.WriteMessage("\nMisc:");
                        ed.WriteMessage($"  Layer: {oleFrame.Layer}");
                        ed.WriteMessage($"  Linetype: {oleFrame.Linetype}");
                        ed.WriteMessage($"  Color Index: {oleFrame.ColorIndex}");
                        ed.WriteMessage($"  Object ID: {oleFrame.ObjectId}");
                        ed.WriteMessage($"  Type: {oleFrame.GetType().Name}");
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
    }
}