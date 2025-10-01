using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.IO;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("GetPdfPageNumber", CommandFlags.Modal)]
        public static void GetPdfPageNumber()
        {
            Document doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Set up selection options to filter for only PDF underlays
            var selOpts = new PromptSelectionOptions();
            selOpts.MessageForAdding = "\nSelect a PDF underlay reference:";
            selOpts.SingleOnly = true;

            var filter = new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "PDFUNDERLAY")
            });

            PromptSelectionResult selRes = ed.GetSelection(selOpts, filter);

            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("*Cancel*");
                return;
            }

            SelectionSet selSet = selRes.Value;
            ObjectId selectedId = selSet.GetObjectIds()[0];

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get the selected PDF underlay reference
                    var pdfRef = tr.GetObject(selectedId, OpenMode.ForRead) as UnderlayReference;
                    if (pdfRef == null)
                    {
                        ed.WriteMessage("\nSelected object is not a valid PDF underlay.");
                        return;
                    }

                    // Get the definition associated with the reference
                    var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                    if (pdfDef == null)
                    {
                        ed.WriteMessage("\nCould not retrieve the PDF definition for the selected reference.");
                        return;
                    }

                    // The 'ItemName' property of the definition holds the page number for PDFs.
                    string pageNumber = pdfDef.ItemName;
                    string pdfFileName = Path.GetFileName(pdfDef.SourceFileName);

                    if (string.IsNullOrEmpty(pageNumber))
                    {
                        ed.WriteMessage($"\nThe selected PDF reference ('{pdfFileName}') does not have a specific page number associated with it.");
                    }
                    else
                    {
                        ed.WriteMessage($"\nThe selected PDF reference is on page: {pageNumber} of '{pdfFileName}'.");
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