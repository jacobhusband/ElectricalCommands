using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

// Alias to avoid ambiguity with System.Windows.Forms.Application
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("IPDF")]
    public void IPDF()
    {
      var doc = AcadApp.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      var db = doc.Database;
      var ed = doc.Editor;

      // 1) Pick the PDF
      OpenFileDialog ofd = new OpenFileDialog
      {
        Filter = "PDF Files (*.pdf)|*.pdf",
        Title = "Select a multi-page PDF to attach"
      };
      if (ofd.ShowDialog() != DialogResult.OK) return;

      string pdfPath = ofd.FileName;
      string baseName = Path.GetFileNameWithoutExtension(pdfPath);

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        // 2) Get or create the PDF underlay definition dictionary
        DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
        const string PDF_DICT_NAME = "ACAD_PDFDEFINITIONS";

        ObjectId pdfDictId = ObjectId.Null;
        DBDictionary pdfDict;

        if (nod.Contains(PDF_DICT_NAME))
        {
          pdfDictId = nod.GetAt(PDF_DICT_NAME);
          pdfDict = (DBDictionary)tr.GetObject(pdfDictId, OpenMode.ForRead);
        }
        else
        {
          nod.UpgradeOpen();
          pdfDict = new DBDictionary();
          pdfDictId = nod.SetAt(PDF_DICT_NAME, pdfDict);
          tr.AddNewlyCreatedDBObject(pdfDict, true);
        }

        // 3) Target Paper Space
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        BlockTableRecord ps = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        // 4) One pick: top-right anchor (to match your T24 placement math)
        PromptPointResult ppr = ed.GetPoint("\nSelect top-right point to insert PDF pages: ");
        if (ppr.Status != PromptStatus.OK) return;
        Point3d anchorTR = ppr.Value;

        // 5) Scale so 1 PDF inch maps to drawing units (true size on layout)
        double scalePerInch = InchesToDrawingUnits(db.Insunits);
        // If you want the PNG-like 80% shrink, uncomment:
        // scalePerInch *= 0.8;

        // Grid controls
        int perRow = 3;
        int col = 0, row = 0;
        Vector3d pageW = Vector3d.XAxis, pageH = Vector3d.YAxis;
        bool firstPlaced = false;

        // 6) Try pages 1..N until a page fails (no dependency on PDF libs)
        const int MAX_PAGES_TRY = 500;
        for (int page = 1; page <= MAX_PAGES_TRY; page++)
        {
          // Unique key for the definition
          string defKey = $"{baseName} - {page}";

          // Create the PdfDefinition for this page
          var pdfDef = new PdfDefinition
          {
            SourceFileName = pdfPath,
            // For PDFs the "ItemName" is the page number as text
            ItemName = page.ToString()
          };

          // Add definition to dictionary
          pdfDict.UpgradeOpen();
          ObjectId defId = pdfDict.SetAt(defKey, pdfDef);
          tr.AddNewlyCreatedDBObject(pdfDef, true);

          // Create the reference and hook the definition
          var pref = new PdfReference
          {
            DefinitionId = defId,          // <-- correct property
            Position = Point3d.Origin,     // temp; we’ll relocate
            Rotation = 0.0,
            Normal = Vector3d.ZAxis,
            ScaleFactors = new Scale3d(scalePerInch)
          };

          try
          {
            ps.AppendEntity(pref);
            tr.AddNewlyCreatedDBObject(pref, true);

            // First page: measure its placed size to drive spacing
            if (!firstPlaced)
            {
              Extents3d e = pref.GeometricExtents;
              double w = Math.Abs(e.MaxPoint.X - e.MinPoint.X);
              double h = Math.Abs(e.MaxPoint.Y - e.MinPoint.Y);
              pageW = new Vector3d(w, 0, 0);
              pageH = new Vector3d(0, h, 0);

              // Move first page so its top-right aligns with anchor
              Point3d ll = new Point3d(anchorTR.X - w, anchorTR.Y - h, 0);
              pref.UpgradeOpen();
              pref.Position = ll;
              firstPlaced = true;
            }
            else
            {
              // For others: LL = anchorTR - ((col+1)*W, (row+1)*H)
              Point3d ll = new Point3d(
                anchorTR.X - ((col + 1) * pageW.X),
                anchorTR.Y - ((row + 1) * pageH.Y),
                0);
              pref.UpgradeOpen();
              pref.Position = ll;
            }

            // Advance 3-wide grid
            col++;
            if (col % perRow == 0)
            {
              row++;
              col = 0;
            }
          }
          catch
          {
            // Page out of range or other failure — clean up def & exit
            try
            {
              pdfDict.UpgradeOpen();
              if (pdfDict.Contains(defKey)) pdfDict.Remove(defKey);
              pdfDef.Dispose();
            }
            catch { /* ignore cleanup issues */ }

            if (!firstPlaced)
              ed.WriteMessage($"\nNo PDF pages could be attached from: {pdfPath}");
            break;
          }
        }

        tr.Commit();
      }
    }

    private static double InchesToDrawingUnits(UnitsValue u)
    {
      // Map drawing INSUNITS to "drawing units per inch"
      switch (u)
      {
        case UnitsValue.Inches:       return 1.0;       // 1 unit = 1 inch
        case UnitsValue.Feet:         return 1.0 / 12.0;
        case UnitsValue.Millimeters:  return 25.4;
        case UnitsValue.Centimeters:  return 2.54;
        case UnitsValue.Meters:       return 0.0254;
        default:                      return 1.0;       // treat as inches
      }
    }
  }
}
