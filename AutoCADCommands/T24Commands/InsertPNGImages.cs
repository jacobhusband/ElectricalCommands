using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("INSERTPNGIMAGES")]
    public void T24()
    {
      Database acCurDb;
      acCurDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
      Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        // Get the user to select a PNG file
        var ofd = new Microsoft.Win32.OpenFileDialog
        {
          Filter = "PNG Files (*.png)|*.png",
          Title = "Select a PNG file"
        };
        if (ofd.ShowDialog() != true)
          return;

        string strFileName = ofd.FileName;

        // Determine the parent folder of the selected file
        string parentFolder = Path.GetDirectoryName(strFileName);

        // Fetch all relevant files in the folder
        string[] allFiles = Directory
            .GetFiles(parentFolder, "*.png")
            .OrderBy(f =>
            {
              var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(f);
              if (fileNameWithoutExtension.Contains("Page"))
              {
                // If the filename contains "Page", we extract the number after it.
                var lastPart = fileNameWithoutExtension.Split(' ').Last();
                return int.Parse(lastPart);
              }
              else
              {
                // If the filename does not contain "Page", we extract the last number after the hyphen.
                var lastPart = fileNameWithoutExtension.Split('-').Last();
                return int.Parse(lastPart);
              }
            })
            .ToArray();

        // Variable to track current image position
        int currentRow = 0;
        int currentColumn = 0;
        Point3d selectedPoint = Point3d.Origin;
        Vector3d width = new Vector3d(0, 0, 0);
        Vector3d height = new Vector3d(0, 0, 0);

        foreach (string file in allFiles)
        {
          string imageName = Path.GetFileNameWithoutExtension(file);

          RasterImageDef acRasterDef;
          bool bRasterDefCreated = false;
          ObjectId acImgDefId;

          // Get the image dictionary
          ObjectId acImgDctID = RasterImageDef.GetImageDictionary(acCurDb);

          // Check to see if the dictionary does not exist, it not then create it
          if (acImgDctID.IsNull)
          {
            acImgDctID = RasterImageDef.CreateImageDictionary(acCurDb);
          }

          // Open the image dictionary
          DBDictionary acImgDict =
              acTrans.GetObject(acImgDctID, OpenMode.ForRead) as DBDictionary;

          // Check to see if the image definition already exists
          if (acImgDict.Contains(imageName))
          {
            acImgDefId = acImgDict.GetAt(imageName);

            acRasterDef =
                acTrans.GetObject(acImgDefId, OpenMode.ForWrite) as RasterImageDef;
          }
          else
          {
            // Create a raster image definition
            RasterImageDef acRasterDefNew = new RasterImageDef();

            // Set the source for the image file
            acRasterDefNew.SourceFileName = file;

            // Load the image into memory
            acRasterDefNew.Load();

            // Add the image definition to the dictionary
            acTrans.GetObject(acImgDctID, OpenMode.ForWrite);
            acImgDefId = acImgDict.SetAt(imageName, acRasterDefNew);

            acTrans.AddNewlyCreatedDBObject(acRasterDefNew, true);

            acRasterDef = acRasterDefNew;

            bRasterDefCreated = true;
          }

          // Open the Block table for read
          BlockTable acBlkTbl;
          acBlkTbl =
              acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

          // Open the Block table record Paper space for write
          BlockTableRecord acBlkTblRec;
          acBlkTblRec =
              acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite)
              as BlockTableRecord;

          // Create the new image and assign it the image definition
          using (RasterImage acRaster = new RasterImage())
          {
            acRaster.ImageDefId = acImgDefId;

            // Define the width and height of the image
            if (selectedPoint == Point3d.Origin)
            {
              // Check to see if the measurement is set to English (Imperial) or Metric units
              if (acCurDb.Measurement == MeasurementValue.English)
              {
                width = new Vector3d(
                    (acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth * 0.8)
                        / 25.4,
                    0,
                    0
                );
                height = new Vector3d(
                    0,
                    (
                        acRasterDef.ResolutionMMPerPixel.Y
                        * acRaster.ImageHeight
                        * 0.8
                    ) / 25.4,
                    0
                );
              }
              else
              {
                width = new Vector3d(
                    acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth * 0.8,
                    0,
                    0
                );
                height = new Vector3d(
                    0,
                    acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight * 0.8,
                    0
                );
              }
            }

            // Prompt the user to select a point
            // Only for the first image
            if (selectedPoint == Point3d.Origin)
            {
              PromptPointResult ppr = ed.GetPoint(
                  "\nSelect a point to insert images:"
              );
              if (ppr.Status != PromptStatus.OK)
                return;
              selectedPoint = ppr.Value;
            }

            // Calculate the new position based on the row and column
            Point3d currentPos = new Point3d(
                selectedPoint.X - (currentColumn * width.X) - width.X, // Subtract width.X to shift the starting point to the top right corner of the first image
                selectedPoint.Y - (currentRow * height.Y) - height.Y, // Subtract height.Y to shift the starting point to the top right corner of the first image
                0
            );

            // Define and assign a coordinate system for the image's orientation
            CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(
                currentPos,
                width,
                height
            );
            acRaster.Orientation = coordinateSystem;

            // Set the rotation angle for the image
            acRaster.Rotation = 0;

            // Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acRaster);
            acTrans.AddNewlyCreatedDBObject(acRaster, true);

            // Connect the raster definition and image together so the definition
            // does not appear as "unreferenced" in the External References palette.
            RasterImage.EnableReactors(true);
            acRaster.AssociateRasterDef(acRasterDef);

            if (bRasterDefCreated)
            {
              acRasterDef.Dispose();
            }
          }

          // Move to the next column
          currentColumn++;

          // Start a new row every 3 images
          if (currentColumn % 3 == 0)
          {
            currentRow++;
            currentColumn = 0;
          }
        }

        // Save the new object to the database
        acTrans.Commit();
      }
    }
  }
}
