using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;
using System.Linq;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        private static Point2d? GetPdfPageSizeInches(string pdfFilePath, int pageNumber)
        {
            if (string.IsNullOrEmpty(pdfFilePath) || !File.Exists(pdfFilePath) || pageNumber <= 0)
            {
                return null;
            }

            try
            {
                using (var pdf = new PdfDocument())
                {
                    pdf.LoadFromFile(pdfFilePath);

                    if (pageNumber > pdf.Pages.Count)
                    {
                        return null; // Page number is out of range
                    }

                    var page = pdf.Pages[pageNumber - 1]; // Spire.Pdf is 0-indexed
                    var sizeInPoints = page.Size;

                    var unitConverter = new PdfUnitConvertor();
                    float widthInches = unitConverter.ConvertUnits(sizeInPoints.Width, PdfGraphicsUnit.Point, PdfGraphicsUnit.Inch);
                    float heightInches = unitConverter.ConvertUnits(sizeInPoints.Height, PdfGraphicsUnit.Point, PdfGraphicsUnit.Inch);

                    return new Point2d(widthInches, heightInches);
                }
            }
            catch (System.Exception)
            {
                // Log exception here if needed
                return null;
            }
        }
        
        [CommandMethod("GETPDFSHEETSIZE", CommandFlags.Modal)]
        public static void GetPdfSheetSize()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect a PDF underlay reference:",
                SingleOnly = true
            };
            var filter = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "PDFUNDERLAY") });

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
                    var ur = tr.GetObject(id, OpenMode.ForRead) as UnderlayReference;
                    if (ur == null)
                    {
                        ed.WriteMessage("\nSelected object is not a PDF underlay.");
                        return;
                    }

                    var def = tr.GetObject(ur.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                    if (def == null)
                    {
                        ed.WriteMessage("\nCould not retrieve PDF definition.");
                        return;
                    }

                    string resolvedPath = ResolveImagePath(db, def.SourceFileName);
                    if (string.IsNullOrEmpty(resolvedPath) || !File.Exists(resolvedPath))
                    {
                        ed.WriteMessage($"\nCould not find PDF file: {def.SourceFileName}");
                        return;
                    }

                    int pageNumber = 1;
                    if (int.TryParse(def.ItemName, out int pg) && pg > 0)
                    {
                        pageNumber = pg;
                    }

                    try
                    {
                        Point2d? pageSize = GetPdfPageSizeInches(resolvedPath, pageNumber);
                        if (pageSize.HasValue)
                        {
                            double wIn = pageSize.Value.X;
                            double hIn = pageSize.Value.Y;
                            string orient = wIn >= hIn ? "Landscape" : "Portrait";
                            string file = Path.GetFileName(def.SourceFileName);

                            ed.WriteMessage($"\nPDF sheet size (inches): {wIn:F2}\" × {hIn:F2}\"  [{orient}] (Page {pageNumber})");
                            ed.WriteMessage($"\nSource: {file}");
                        }
                        else
                        {
                            ed.WriteMessage("\nCould not determine PDF page size using Spire.PDF.");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError while getting PDF size: {ex.Message}");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }
    }
}
