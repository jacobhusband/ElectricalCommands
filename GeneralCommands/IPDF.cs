using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

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

      OpenFileDialog ofd = new OpenFileDialog
      {
        Filter = "PDF Files (*.pdf)|*.pdf",
        Title = "Select a multi-page PDF to attach"
      };
      if (ofd.ShowDialog() != DialogResult.OK) return;

      string pdfPath = ofd.FileName;
      string baseName = Path.GetFileNameWithoutExtension(pdfPath);
      string defPrefix = $"{baseName}_{ComputeStableHashSuffix(pdfPath)}";

      if (!TryGetPdfPageCount(pdfPath, out int pageCount))
      {
        ed.WriteMessage("\nUnable to determine PDF page count automatically.");
        var pagePrompt = new PromptIntegerOptions("\nEnter the number of PDF pages to insert")
        {
          AllowZero = false,
          AllowNegative = false,
          LowerLimit = 1,
          DefaultValue = 1
        };
        var countResult = ed.GetInteger(pagePrompt);
        if (countResult.Status != PromptStatus.OK) return;
        pageCount = countResult.Value;
      }

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
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

        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        BlockTableRecord ps = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        PromptPointResult ppr = ed.GetPoint("\nSelect top-right point to insert PDF pages: ");
        if (ppr.Status != PromptStatus.OK) return;
        Point3d anchorTR = ppr.Value;

        double scalePerInch = InchesToDrawingUnits(db.Insunits);

        int perRow = 3;
        int col = 0, row = 0;
        Vector3d pageW = Vector3d.XAxis, pageH = Vector3d.YAxis;
        bool firstPlaced = false;

        for (int page = 1; page <= pageCount; page++)
        {
          string defKey = $"{defPrefix} - {page}";
          ObjectId defId;
          bool definitionCreated = false;

          if (pdfDict.Contains(defKey))
          {
            defId = pdfDict.GetAt(defKey);
          }
          else
          {
            if (!pdfDict.IsWriteEnabled)
            {
              pdfDict.UpgradeOpen();
            }

            var pdfDef = new PdfDefinition
            {
              SourceFileName = pdfPath,
              ItemName = page.ToString()
            };

            defId = pdfDict.SetAt(defKey, pdfDef);
            tr.AddNewlyCreatedDBObject(pdfDef, true);
            definitionCreated = true;
          }

          var pref = new PdfReference
          {
            DefinitionId = defId,
            Position = Point3d.Origin,
            Rotation = 0.0,
            Normal = Vector3d.ZAxis,
            ScaleFactors = new Scale3d(scalePerInch)
          };

          try
          {
            ps.AppendEntity(pref);
            tr.AddNewlyCreatedDBObject(pref, true);

            if (!firstPlaced)
            {
              Extents3d extents = pref.GeometricExtents;
              double w = Math.Abs(extents.MaxPoint.X - extents.MinPoint.X);
              double h = Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y);

              if (w <= Tolerance.Global.EqualPoint || h <= Tolerance.Global.EqualPoint)
              {
                throw new InvalidOperationException("Unable to determine PDF page dimensions.");
              }

              pageW = new Vector3d(w, 0, 0);
              pageH = new Vector3d(0, h, 0);

              pref.UpgradeOpen();
              pref.Position = new Point3d(anchorTR.X - w, anchorTR.Y - h, 0);
              pref.DowngradeOpen();

              firstPlaced = true;
            }
            else
            {
              Point3d ll = new Point3d(
                anchorTR.X - ((col + 1) * pageW.X),
                anchorTR.Y - ((row + 1) * pageH.Y),
                0);

              pref.UpgradeOpen();
              pref.Position = ll;
              pref.DowngradeOpen();
            }

            col++;
            if (col % perRow == 0)
            {
              row++;
              col = 0;
            }
          }
          catch (System.Exception ex)
          {
            ed.WriteMessage($"\nStopped inserting at page {page}: {ex.Message}");
            TryEraseReference(pref);

            if (definitionCreated)
            {
              TryRemoveDefinition(pdfDict, defKey);
            }

            break;
          }
        }

        tr.Commit();
      }
    }

    private static bool TryGetPdfPageCount(string pdfPath, out int pageCount)
    {
      try
      {
        const int ReadLimitBytes = 4 * 1024 * 1024;

        using (FileStream fs = new FileStream(pdfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
          int bufferSize = (int)Math.Min(ReadLimitBytes, fs.Length);
          if (bufferSize <= 0)
          {
            pageCount = 0;
            return false;
          }

          byte[] buffer = new byte[bufferSize];
          int bytesRead = fs.Read(buffer, 0, bufferSize);
          if (bytesRead <= 0)
          {
            pageCount = 0;
            return false;
          }

          string text = Encoding.ASCII.GetString(buffer, 0, bytesRead);
          int maxCount = 0;

          foreach (Match match in Regex.Matches(text, @"/Type\s*/Pages\b[^>]*?/Count\s+(\d+)", RegexOptions.IgnoreCase | RegexOptions.Singleline))
          {
            if (int.TryParse(match.Groups[1].Value, out int candidate) && candidate > maxCount)
            {
              maxCount = candidate;
            }
          }

          if (maxCount == 0)
          {
            foreach (Match match in Regex.Matches(text, @"/Count\s+(\d+)", RegexOptions.IgnoreCase))
            {
              if (int.TryParse(match.Groups[1].Value, out int candidate) && candidate > maxCount)
              {
                maxCount = candidate;
              }
            }
          }

          if (maxCount > 0)
          {
            pageCount = maxCount;
            return true;
          }
        }
      }
      catch
      {
      }

      pageCount = 0;
      return false;
    }
    private static void TryEraseReference(PdfReference reference)
    {
      if (reference == null) return;

      try
      {
        if (reference.ObjectId.IsNull) return;
        if (!reference.IsWriteEnabled) reference.UpgradeOpen();
        reference.Erase();
      }
      catch
      {
        // Ignore cleanup issues during error handling.
      }
    }

    private static void TryRemoveDefinition(DBDictionary pdfDict, string defKey)
    {
      if (pdfDict == null) return;

      try
      {
        if (!pdfDict.IsWriteEnabled) pdfDict.UpgradeOpen();
        if (pdfDict.Contains(defKey))
        {
          pdfDict.Remove(defKey);
        }
      }
      catch
      {
        // Ignore cleanup issues.
      }
    }

    private static string ComputeStableHashSuffix(string input)
    {
      using var sha1 = SHA1.Create();
      byte[] hash = sha1.ComputeHash(Encoding.UTF8.GetBytes(input.ToLowerInvariant()));
      return BitConverter.ToString(hash, 0, 3).Replace("-", string.Empty);
    }

    private static double InchesToDrawingUnits(UnitsValue u)
    {
      switch (u)
      {
        case UnitsValue.Inches:       return 1.0;
        case UnitsValue.Feet:         return 1.0 / 12.0;
        case UnitsValue.Millimeters:  return 25.4;
        case UnitsValue.Centimeters:  return 2.54;
        case UnitsValue.Meters:       return 0.0254;
        default:                      return 1.0;
      }
    }
  }
}




