using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        /// <summary>
        /// A diagnostic command to identify the type of any selected entity.
        /// </summary>
        [CommandMethod("GETOBJECTINFO", CommandFlags.Modal)]
        public static void GetObjectInfo()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect any object to identify its type: ");
            var per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nOperation cancelled.");
                return;
            }

            // The ObjectId is available directly from the prompt result
            ObjectId selectedId = per.ObjectId;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // We still open the object to get its full .NET type name
                    var selectedObject = tr.GetObject(selectedId, OpenMode.ForRead);

                    ed.WriteMessage("\n--- Selected Object Information ---");

                    // CORRECTED: Get the ObjectClass and DxfName from the ObjectId.
                    // This is the correct and compatible way.
                    ed.WriteMessage($"\n  DXF Type Name: {selectedId.ObjectClass.DxfName}");

                    // The .NET type is the name of the class in the API
                    ed.WriteMessage($"\n  .NET Class Name: {selectedObject.GetType().FullName}");
                    ed.WriteMessage("\n-----------------------------------");
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