using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ElectricalCommands
{
  /// <summary>
  /// Data model for storing text object properties relative to a base point.
  /// </summary>
  public class TextInfoData
  {
    public string TextType { get; set; } // "DBText" or "MText"
    public string TextContent { get; set; }
    public double OffsetX { get; set; } // Offset from base point
    public double OffsetY { get; set; }
    public double OffsetZ { get; set; }
    public double Height { get; set; }
    public double Rotation { get; set; }
    public string TextStyleName { get; set; }
    public string Layer { get; set; }
    public int ColorIndex { get; set; }

    // DBText specific properties
    public double WidthFactor { get; set; }
    public double Oblique { get; set; }
    public int HorizontalMode { get; set; }
    public int VerticalMode { get; set; }
    public double AlignmentOffsetX { get; set; }
    public double AlignmentOffsetY { get; set; }
    public double AlignmentOffsetZ { get; set; }

    // MText specific properties
    public double MTextWidth { get; set; }
    public int Attachment { get; set; }
  }

  /// <summary>
  /// Data model for storing raster image properties relative to a base point.
  /// </summary>
  public class ImageInfoData
  {
    public string ImageFileName { get; set; } // Just the filename, not full path
    public double OffsetX { get; set; }
    public double OffsetY { get; set; }
    public double OffsetZ { get; set; }
    public double ScaleX { get; set; }
    public double ScaleY { get; set; }
    public double Rotation { get; set; }
    public string Layer { get; set; }
  }

  /// <summary>
  /// Container for the saved text info collection.
  /// </summary>
  public class TextInfoCollection
  {
    public List<TextInfoData> TextObjects { get; set; } = new List<TextInfoData>();
    public DateTime SavedDate { get; set; }
    public string SourceDrawing { get; set; }
  }

  /// <summary>
  /// Container for the saved image info collection.
  /// </summary>
  public class ImageInfoCollection
  {
    public List<ImageInfoData> Images { get; set; } = new List<ImageInfoData>();
    public DateTime SavedDate { get; set; }
    public string SourceDrawing { get; set; }
  }

  public partial class GeneralCommands
  {
    // ============================================================================
    // EMBEDDED T24 FORM DATA - Hardcoded text and image information
    // ============================================================================

    /// <summary>
    /// Embedded T24 text data for automatic form filling.
    /// These are the constant text entries that will be placed on the T24 form.
    /// </summary>
    private static readonly List<TextInfoData> T24TextData = new List<TextInfoData>
    {
      // Row 1 - Left column (Responsible Person info)
      new TextInfoData { TextType = "MText", TextContent = "Wilson Lee, PE", OffsetX = 2.4575724590123826, OffsetY = -4.605522133300136, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "ACIES Engineering", OffsetX = 2.4575724590123826, OffsetY = -4.78756234623763, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "400 N McCarthy Blvd., Suite 250", OffsetX = 2.4575724590123826, OffsetY = -4.968225407280139, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "Milipitas, CA 95035", OffsetX = 2.4575724590123826, OffsetY = -5.141314963399362, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },

      // Row 1 - Right column (Phone, License, Date)
      new TextInfoData { TextType = "MText", TextContent = "(408) 522-5255", OffsetX = 7.015472016179885, OffsetY = -5.140839890951543, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "E015418", OffsetX = 7.067721730676951, OffsetY = -4.962215889505831, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "{DATE}", OffsetX = 7.0577125306994475, OffsetY = -4.790458176947315, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },

      // Row 2 - Left column (Responsible Person info - duplicate section)
      new TextInfoData { TextType = "MText", TextContent = "Wilson Lee, PE", OffsetX = 2.4575724590123826, OffsetY = -2.177877868536209, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "ACIES Engineering", OffsetX = 2.4575724590123826, OffsetY = -2.489218470561264, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "400 N McCarthy Blvd., Suite 250", OffsetX = 2.4575724590123826, OffsetY = -2.7560817909954896, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "Milipitas, CA 95035", OffsetX = 2.4575724590123826, OffsetY = -2.93997586225313, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },

      // Row 2 - Right column (Phone, Date)
      new TextInfoData { TextType = "MText", TextContent = "(408) 522-5255", OffsetX = 6.8598242698728455, OffsetY = -2.9268530282783622, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
      new TextInfoData { TextType = "MText", TextContent = "{DATE}", OffsetX = 6.8598242698728455, OffsetY = -2.5170374839403884, Height = 0.10546875, TextStyleName = "Arial Narrow_1", Layer = "E-TEXT", ColorIndex = 256, Attachment = 1 },
    };

    /// <summary>
    /// Embedded T24 signature image data.
    /// Image files should be in the same folder as the DWG.
    /// </summary>
    private static readonly List<ImageInfoData> T24ImageData = new List<ImageInfoData>
    {
      // Signature image - Row 2 (upper section on form)
      new ImageInfoData { ImageFileName = "WL_Sig_Blue_Small.gif", OffsetX = 7.081875433832849, OffsetY = -2.33079552101281, ScaleX = 0.4332543938729146, ScaleY = 0.16140849967814466, Rotation = 0.0, Layer = "0" },
      // Signature image - Row 1 (lower section on form)
      new ImageInfoData { ImageFileName = "WL_Sig_Blue_Small.gif", OffsetX = 7.081875433832849, OffsetY = -4.744404260423231, ScaleX = 0.4332543938729146, ScaleY = 0.16140849967814466, Rotation = 0.0, Layer = "0" },
    };

    // ============================================================================
    // FILE PATH HELPERS
    // ============================================================================

    private static string GetTextInfoFilePath()
    {
      string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
      string folderPath = Path.Combine(appDataPath, "ElectricalCommands");
      if (!Directory.Exists(folderPath))
      {
        Directory.CreateDirectory(folderPath);
      }
      return Path.Combine(folderPath, "TextInfo.json");
    }

    private static string GetImageInfoFilePath()
    {
      string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
      string folderPath = Path.Combine(appDataPath, "ElectricalCommands");
      if (!Directory.Exists(folderPath))
      {
        Directory.CreateDirectory(folderPath);
      }
      return Path.Combine(folderPath, "ImageInfo.json");
    }

    // ============================================================================
    // TEXTINFO COMMAND - Capture text object positions
    // ============================================================================

    /// <summary>
    /// TEXTINFO command - Select text objects and a base point, save all properties to JSON.
    /// </summary>
    [CommandMethod("TEXTINFO", CommandFlags.UsePickSet)]
    public void TextInfo()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null) return;

      try
      {
        // 1. Select text objects
        PromptSelectionResult psr = ed.SelectImplied();
        if (psr.Status != PromptStatus.OK)
        {
          PromptSelectionOptions pso = new PromptSelectionOptions();
          pso.MessageForAdding = "Select DBText or MText objects: ";
          TypedValue[] filterList = new TypedValue[]
          {
            new TypedValue((int)DxfCode.Start, "TEXT,MTEXT")
          };
          SelectionFilter filter = new SelectionFilter(filterList);
          psr = ed.GetSelection(pso, filter);
        }

        if (psr.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo objects selected.");
          return;
        }

        // 2. Get base point
        PromptPointResult ppr = ed.GetPoint("\nSelect base point: ");
        if (ppr.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo base point selected.");
          return;
        }
        Point3d basePoint = ppr.Value;

        // 3. Collect text info
        TextInfoCollection collection = new TextInfoCollection
        {
          SavedDate = DateTime.Now,
          SourceDrawing = db.Filename
        };

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          SelectionSet ss = psr.Value;
          foreach (SelectedObject so in ss)
          {
            Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
            if (ent == null) continue;

            TextInfoData info = new TextInfoData();

            if (ent is DBText dbText)
            {
              info.TextType = "DBText";
              info.TextContent = dbText.TextString;
              info.OffsetX = dbText.Position.X - basePoint.X;
              info.OffsetY = dbText.Position.Y - basePoint.Y;
              info.OffsetZ = dbText.Position.Z - basePoint.Z;
              info.Height = dbText.Height;
              info.Rotation = dbText.Rotation;
              info.Layer = dbText.Layer;
              info.ColorIndex = dbText.ColorIndex;
              info.WidthFactor = dbText.WidthFactor;
              info.Oblique = dbText.Oblique;
              info.HorizontalMode = (int)dbText.HorizontalMode;
              info.VerticalMode = (int)dbText.VerticalMode;
              info.AlignmentOffsetX = dbText.AlignmentPoint.X - basePoint.X;
              info.AlignmentOffsetY = dbText.AlignmentPoint.Y - basePoint.Y;
              info.AlignmentOffsetZ = dbText.AlignmentPoint.Z - basePoint.Z;

              // Get text style name
              TextStyleTable tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
              TextStyleTableRecord tstr = (TextStyleTableRecord)tr.GetObject(dbText.TextStyleId, OpenMode.ForRead);
              info.TextStyleName = tstr.Name;
            }
            else if (ent is MText mText)
            {
              info.TextType = "MText";
              info.TextContent = mText.Contents;
              info.OffsetX = mText.Location.X - basePoint.X;
              info.OffsetY = mText.Location.Y - basePoint.Y;
              info.OffsetZ = mText.Location.Z - basePoint.Z;
              info.Height = mText.TextHeight;
              info.Rotation = mText.Rotation;
              info.Layer = mText.Layer;
              info.ColorIndex = mText.ColorIndex;
              info.MTextWidth = mText.Width;
              info.Attachment = (int)mText.Attachment;

              // Get text style name
              TextStyleTable tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
              TextStyleTableRecord tstr = (TextStyleTableRecord)tr.GetObject(mText.TextStyleId, OpenMode.ForRead);
              info.TextStyleName = tstr.Name;
            }
            else
            {
              continue;
            }

            collection.TextObjects.Add(info);
          }

          tr.Commit();
        }

        // 4. Save to JSON
        string jsonPath = GetTextInfoFilePath();
        string json = JsonConvert.SerializeObject(collection, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        ed.WriteMessage($"\nSaved {collection.TextObjects.Count} text object(s) to: {jsonPath}");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nError: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    // ============================================================================
    // IMAGEINFO COMMAND - Capture raster image positions
    // ============================================================================

    /// <summary>
    /// IMAGEINFO command - Select raster images and a base point, save all properties to JSON.
    /// </summary>
    [CommandMethod("IMAGEINFO", CommandFlags.UsePickSet)]
    public void ImageInfo()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null) return;

      try
      {
        // 1. Select raster images
        PromptSelectionResult psr = ed.SelectImplied();
        if (psr.Status != PromptStatus.OK)
        {
          PromptSelectionOptions pso = new PromptSelectionOptions();
          pso.MessageForAdding = "Select raster images: ";
          TypedValue[] filterList = new TypedValue[]
          {
            new TypedValue((int)DxfCode.Start, "IMAGE")
          };
          SelectionFilter filter = new SelectionFilter(filterList);
          psr = ed.GetSelection(pso, filter);
        }

        if (psr.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo objects selected.");
          return;
        }

        // 2. Get base point
        PromptPointResult ppr = ed.GetPoint("\nSelect base point: ");
        if (ppr.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo base point selected.");
          return;
        }
        Point3d basePoint = ppr.Value;

        // 3. Collect image info
        ImageInfoCollection collection = new ImageInfoCollection
        {
          SavedDate = DateTime.Now,
          SourceDrawing = db.Filename
        };

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          // Get image dictionary
          DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
          DBDictionary imgDict = null;
          if (nod.Contains("ACAD_IMAGE_DICT"))
          {
            imgDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForRead);
          }

          SelectionSet ss = psr.Value;
          foreach (SelectedObject so in ss)
          {
            Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
            if (ent == null) continue;

            if (ent is RasterImage rasterImage)
            {
              ImageInfoData info = new ImageInfoData();

              // Get image definition to find filename
              RasterImageDef imgDef = (RasterImageDef)tr.GetObject(rasterImage.ImageDefId, OpenMode.ForRead);
              info.ImageFileName = Path.GetFileName(imgDef.SourceFileName);

              // Get position (origin of the image)
              Point3d imgOrigin = rasterImage.Orientation.Origin;
              info.OffsetX = imgOrigin.X - basePoint.X;
              info.OffsetY = imgOrigin.Y - basePoint.Y;
              info.OffsetZ = imgOrigin.Z - basePoint.Z;

              // Get scale from orientation vectors
              Vector3d uVec = rasterImage.Orientation.Xaxis;
              Vector3d vVec = rasterImage.Orientation.Yaxis;
              info.ScaleX = uVec.Length;
              info.ScaleY = vVec.Length;

              info.Rotation = rasterImage.Rotation;
              info.Layer = rasterImage.Layer;

              collection.Images.Add(info);
              ed.WriteMessage($"\n  Found image: {info.ImageFileName} at offset ({info.OffsetX:F4}, {info.OffsetY:F4})");
            }
          }

          tr.Commit();
        }

        // 4. Save to JSON
        string jsonPath = GetImageInfoFilePath();
        string json = JsonConvert.SerializeObject(collection, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        ed.WriteMessage($"\n\nSaved {collection.Images.Count} image(s) to: {jsonPath}");
        ed.WriteMessage("\n\nCopy the image data to the T24ImageData list in the source code to embed it.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nError: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    // ============================================================================
    // CREATETEXTFROMINFO COMMAND - Recreate text from saved JSON
    // ============================================================================

    /// <summary>
    /// CREATETEXTFROMINFO command - Load saved text info and recreate text at a new base point.
    /// </summary>
    [CommandMethod("CREATETEXTFROMINFO")]
    public void CreateTextFromInfo()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null) return;

      try
      {
        // 1. Load JSON
        string jsonPath = GetTextInfoFilePath();
        if (!File.Exists(jsonPath))
        {
          ed.WriteMessage("\nNo saved text info found. Run TEXTINFO first to save text objects.");
          return;
        }

        string json = File.ReadAllText(jsonPath);
        TextInfoCollection collection = JsonConvert.DeserializeObject<TextInfoCollection>(json);

        if (collection == null || collection.TextObjects.Count == 0)
        {
          ed.WriteMessage("\nNo text objects found in saved data.");
          return;
        }

        ed.WriteMessage($"\nLoaded {collection.TextObjects.Count} text object(s) saved on {collection.SavedDate}");

        // 2. Get new base point
        PromptPointResult ppr = ed.GetPoint("\nSelect new base point: ");
        if (ppr.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo base point selected.");
          return;
        }
        Point3d newBasePoint = ppr.Value;

        // 3. Create text objects
        int createdCount = CreateTextObjectsAtPoint(db, ed, collection.TextObjects, newBasePoint, true);
        ed.WriteMessage($"\nCreated {createdCount} text object(s).");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nError: {ex.Message}");
      }
    }

    // ============================================================================
    // T24 COMMAND - Automated PDF insertion with text/image placement
    // ============================================================================

    /// <summary>
    /// T24 command - Insert PDF sheets and automatically add text and images on the last page.
    /// </summary>
    [CommandMethod("T24")]
    public void T24Command()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null) return;

      try
      {
        // 1. Open file dialog for PDF selection
        var ofd = new Microsoft.Win32.OpenFileDialog
        {
          Filter = "PDF Files (*.pdf)|*.pdf",
          Title = "Select T24 PDF form to attach"
        };
        if (ofd.ShowDialog() != true) return;

        string pdfPath = ofd.FileName;
        string baseName = Path.GetFileNameWithoutExtension(pdfPath);
        string defPrefix = $"{baseName}_{ComputeStableHashSuffixT24(pdfPath)}";

        // 2. Get PDF page count
        if (!TryGetPdfPageCountT24(pdfPath, out int pageCount))
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

        ed.WriteMessage($"\nInserting {pageCount} PDF page(s)...");

        // 3. Get anchor point
        PromptPointResult ppr = ed.GetPoint("\nSelect top-right point to insert PDF pages: ");
        if (ppr.Status != PromptStatus.OK) return;
        Point3d anchorTR = ppr.Value;

        // Track the last page's top-left corner
        Point3d lastPageTopLeft = Point3d.Origin;
        int successfulPages = 0;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
          const string PDF_DICT_NAME = "ACAD_PDFDEFINITIONS";

          DBDictionary pdfDict;
          if (nod.Contains(PDF_DICT_NAME))
          {
            pdfDict = (DBDictionary)tr.GetObject(nod.GetAt(PDF_DICT_NAME), OpenMode.ForRead);
          }
          else
          {
            nod.UpgradeOpen();
            pdfDict = new DBDictionary();
            nod.SetAt(PDF_DICT_NAME, pdfDict);
            tr.AddNewlyCreatedDBObject(pdfDict, true);
          }

          BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
          BlockTableRecord ps = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

          double scalePerInch = InchesToDrawingUnitsT24(db.Insunits);

          int perRow = 3;
          int col = 0, row = 0;
          Vector3d pageW = Vector3d.XAxis, pageH = Vector3d.YAxis;
          bool firstPlaced = false;

          // Loop through pages
          for (int page = 1; page <= pageCount; page++)
          {
            string defKey = $"{defPrefix} - {page}";
            ObjectId defId;
            bool definitionCreated = false;

            if (!pdfDict.Contains(defKey))
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
            else
            {
              defId = pdfDict.GetAt(defKey);
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

              Extents3d extents = pref.GeometricExtents;
              double w = Math.Abs(extents.MaxPoint.X - extents.MinPoint.X);
              double h = Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y);

              if (w <= Tolerance.Global.EqualPoint || h <= Tolerance.Global.EqualPoint)
              {
                ed.WriteMessage($"\nStopped at page {page}. Assumed end of document.");
                TryEraseReferenceT24(pref);
                if (definitionCreated)
                {
                  TryRemoveDefinitionT24(pdfDict, defKey);
                }
                break;
              }

              Point3d pagePosition;

              if (!firstPlaced)
              {
                pageW = new Vector3d(w, 0, 0);
                pageH = new Vector3d(0, h, 0);

                pagePosition = new Point3d(anchorTR.X - w, anchorTR.Y - h, 0);
                pref.UpgradeOpen();
                pref.Position = pagePosition;
                pref.DowngradeOpen();

                firstPlaced = true;
              }
              else
              {
                pagePosition = new Point3d(
                    (anchorTR.X - pageW.X) - (col * pageW.X),
                    (anchorTR.Y - pageH.Y) - (row * pageH.Y),
                    0);

                pref.UpgradeOpen();
                pref.Position = pagePosition;
                pref.DowngradeOpen();
              }

              // Track the last page's top-left corner
              // Position is bottom-left, so top-left = position + height
              lastPageTopLeft = new Point3d(pagePosition.X, pagePosition.Y + h, 0);

              successfulPages++;
              col++;
              if (col >= perRow)
              {
                row++;
                col = 0;
              }
            }
            catch (System.Exception ex)
            {
              ed.WriteMessage($"\nAn error occurred while inserting page {page}: {ex.Message}");
              TryEraseReferenceT24(pref);
              if (definitionCreated)
              {
                TryRemoveDefinitionT24(pdfDict, defKey);
              }
              break;
            }
          }

          tr.Commit();
        }

        if (successfulPages > 0)
        {
          ed.WriteMessage($"\nSuccessfully inserted {successfulPages} page(s).");

          // 4. Now place text and images on the last page
          ed.WriteMessage($"\nPlacing text on last page at top-left: ({lastPageTopLeft.X:F4}, {lastPageTopLeft.Y:F4})");

          int textCount = CreateTextObjectsAtPoint(db, ed, T24TextData, lastPageTopLeft, true);
          ed.WriteMessage($"\nCreated {textCount} text object(s) with current date: {DateTime.Now:MM/dd/yyyy}");

          // 5. Place images if configured
          if (T24ImageData.Count > 0)
          {
            string dwgFolder = Path.GetDirectoryName(db.Filename);
            int imageCount = CreateImagesAtPoint(db, ed, T24ImageData, lastPageTopLeft, dwgFolder);
            ed.WriteMessage($"\nCreated {imageCount} image(s).");
          }
        }
        else
        {
          ed.WriteMessage("\nNo pages were inserted.");
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nError: {ex.Message}");
      }
    }

    // ============================================================================
    // HELPER METHODS
    // ============================================================================

    /// <summary>
    /// Creates text objects at a specified base point from a list of TextInfoData.
    /// </summary>
    private int CreateTextObjectsAtPoint(Database db, Editor ed, List<TextInfoData> textDataList, Point3d basePoint, bool replaceDatePlaceholder)
    {
      int createdCount = 0;
      string currentDate = DateTime.Now.ToString("MM/dd/yyyy");

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTableRecord currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        TextStyleTable tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

        foreach (TextInfoData info in textDataList)
        {
          try
          {
            // Find or use default text style
            ObjectId textStyleId = db.Textstyle;
            if (tst.Has(info.TextStyleName))
            {
              textStyleId = tst[info.TextStyleName];
            }

            // Replace date placeholder if requested
            string textContent = info.TextContent;
            if (replaceDatePlaceholder && (textContent.Contains("{DATE}") || textContent.Contains("{date}")))
            {
              textContent = textContent.Replace("{DATE}", currentDate).Replace("{date}", currentDate);
            }

            if (info.TextType == "DBText")
            {
              DBText newText = new DBText();
              newText.SetDatabaseDefaults();
              newText.TextString = textContent;
              newText.Position = new Point3d(
                basePoint.X + info.OffsetX,
                basePoint.Y + info.OffsetY,
                basePoint.Z + info.OffsetZ);
              newText.Height = info.Height;
              newText.Rotation = info.Rotation;
              newText.TextStyleId = textStyleId;
              newText.Layer = EnsureLayerExists(tr, db, info.Layer);
              newText.ColorIndex = info.ColorIndex;
              newText.WidthFactor = info.WidthFactor > 0 ? info.WidthFactor : 1.0;
              newText.Oblique = info.Oblique;
              newText.HorizontalMode = (TextHorizontalMode)info.HorizontalMode;
              newText.VerticalMode = (TextVerticalMode)info.VerticalMode;
              newText.AlignmentPoint = new Point3d(
                basePoint.X + info.AlignmentOffsetX,
                basePoint.Y + info.AlignmentOffsetY,
                basePoint.Z + info.AlignmentOffsetZ);

              currentSpace.AppendEntity(newText);
              tr.AddNewlyCreatedDBObject(newText, true);
              createdCount++;
            }
            else if (info.TextType == "MText")
            {
              MText newText = new MText();
              newText.SetDatabaseDefaults();
              newText.Contents = textContent;
              newText.Location = new Point3d(
                basePoint.X + info.OffsetX,
                basePoint.Y + info.OffsetY,
                basePoint.Z + info.OffsetZ);
              newText.TextHeight = info.Height;
              newText.Rotation = info.Rotation;
              newText.TextStyleId = textStyleId;
              newText.Layer = EnsureLayerExists(tr, db, info.Layer);
              newText.ColorIndex = info.ColorIndex;
              newText.Width = info.MTextWidth;
              newText.Attachment = (AttachmentPoint)info.Attachment;

              currentSpace.AppendEntity(newText);
              tr.AddNewlyCreatedDBObject(newText, true);
              createdCount++;
            }
          }
          catch (System.Exception ex)
          {
            ed.WriteMessage($"\nWarning: Could not create text '{info.TextContent}': {ex.Message}");
          }
        }

        tr.Commit();
      }

      return createdCount;
    }

    /// <summary>
    /// Creates raster images at a specified base point from a list of ImageInfoData.
    /// </summary>
    private int CreateImagesAtPoint(Database db, Editor ed, List<ImageInfoData> imageDataList, Point3d basePoint, string imageFolder)
    {
      int createdCount = 0;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        // Get or create image dictionary
        DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
        const string IMG_DICT_NAME = "ACAD_IMAGE_DICT";

        DBDictionary imgDict;
        if (nod.Contains(IMG_DICT_NAME))
        {
          imgDict = (DBDictionary)tr.GetObject(nod.GetAt(IMG_DICT_NAME), OpenMode.ForWrite);
        }
        else
        {
          nod.UpgradeOpen();
          imgDict = new DBDictionary();
          nod.SetAt(IMG_DICT_NAME, imgDict);
          tr.AddNewlyCreatedDBObject(imgDict, true);
        }

        BlockTableRecord currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

        foreach (ImageInfoData info in imageDataList)
        {
          try
          {
            string imagePath = Path.Combine(imageFolder, info.ImageFileName);

            if (!File.Exists(imagePath))
            {
              ed.WriteMessage($"\nWarning: Image file not found: {imagePath}");
              continue;
            }

            // Create or get image definition
            string defKey = info.ImageFileName;
            ObjectId imgDefId;

            if (imgDict.Contains(defKey))
            {
              imgDefId = imgDict.GetAt(defKey);
            }
            else
            {
              RasterImageDef imgDef = new RasterImageDef();
              imgDef.SourceFileName = imagePath;
              imgDef.Load();

              imgDefId = imgDict.SetAt(defKey, imgDef);
              tr.AddNewlyCreatedDBObject(imgDef, true);
            }

            // Create image reference
            RasterImage rasterImage = new RasterImage();
            rasterImage.ImageDefId = imgDefId;

            // Set position and scale using coordinate system
            Point3d position = new Point3d(
              basePoint.X + info.OffsetX,
              basePoint.Y + info.OffsetY,
              basePoint.Z + info.OffsetZ);

            Vector3d uVec = new Vector3d(info.ScaleX, 0, 0);
            Vector3d vVec = new Vector3d(0, info.ScaleY, 0);

            if (info.Rotation != 0)
            {
              double cos = Math.Cos(info.Rotation);
              double sin = Math.Sin(info.Rotation);
              uVec = new Vector3d(info.ScaleX * cos, info.ScaleX * sin, 0);
              vVec = new Vector3d(-info.ScaleY * sin, info.ScaleY * cos, 0);
            }

            rasterImage.Orientation = new CoordinateSystem3d(position, uVec, vVec);
            rasterImage.Layer = EnsureLayerExists(tr, db, info.Layer);

            currentSpace.AppendEntity(rasterImage);
            tr.AddNewlyCreatedDBObject(rasterImage, true);

            createdCount++;
          }
          catch (System.Exception ex)
          {
            ed.WriteMessage($"\nWarning: Could not create image '{info.ImageFileName}': {ex.Message}");
          }
        }

        tr.Commit();
      }

      return createdCount;
    }

    /// <summary>
    /// Helper to ensure a layer exists, returns "0" if the specified layer doesn't exist.
    /// </summary>
    private string EnsureLayerExists(Transaction tr, Database db, string layerName)
    {
      LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
      if (lt.Has(layerName))
      {
        return layerName;
      }
      return "0";
    }

    // PDF helper methods (duplicated from InsertPDFSheets to avoid dependency issues)
    private static bool TryGetPdfPageCountT24(string pdfPath, out int pageCount)
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

    private static void TryEraseReferenceT24(PdfReference reference)
    {
      if (reference == null) return;
      try
      {
        if (reference.ObjectId.IsNull) return;
        if (!reference.IsWriteEnabled) reference.UpgradeOpen();
        reference.Erase();
      }
      catch { }
    }

    private static void TryRemoveDefinitionT24(DBDictionary pdfDict, string defKey)
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
      catch { }
    }

    private static string ComputeStableHashSuffixT24(string input)
    {
      using (var sha1 = SHA1.Create())
      {
        byte[] hash = sha1.ComputeHash(Encoding.UTF8.GetBytes(input.ToLowerInvariant()));
        return BitConverter.ToString(hash, 0, 3).Replace("-", string.Empty);
      }
    }

    private static double InchesToDrawingUnitsT24(UnitsValue u)
    {
      switch (u)
      {
        case UnitsValue.Inches: return 1.0;
        case UnitsValue.Feet: return 1.0 / 12.0;
        case UnitsValue.Millimeters: return 25.4;
        case UnitsValue.Centimeters: return 2.54;
        case UnitsValue.Meters: return 0.0254;
        default: return 1.0;
      }
    }
  }
}
