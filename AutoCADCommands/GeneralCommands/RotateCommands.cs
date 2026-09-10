using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("R90", CommandFlags.UsePickSet)]
    public static void Rotate90()
    {
      RotateByDegrees(90.0);
    }

    [CommandMethod("R180", CommandFlags.UsePickSet)]
    public static void Rotate180()
    {
      RotateByDegrees(180.0);
    }

    [CommandMethod("R270", CommandFlags.UsePickSet)]
    public static void Rotate270()
    {
      RotateByDegrees(270.0);
    }

    [CommandMethod("R360", CommandFlags.UsePickSet)]
    public static void Rotate360()
    {
      RotateByDegrees(360.0);
    }

    private static void RotateByDegrees(double degrees)
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      try
      {
        PromptSelectionResult selectionResult = ed.SelectImplied();

        if (selectionResult.Status != PromptStatus.OK ||
            selectionResult.Value == null ||
            selectionResult.Value.Count == 0)
        {
          PromptSelectionOptions selectionOptions = new PromptSelectionOptions
          {
            MessageForAdding = $"\nSelect objects to rotate {degrees:0} degrees: "
          };

          selectionResult = ed.GetSelection(selectionOptions);
        }

        if (selectionResult.Status != PromptStatus.OK ||
            selectionResult.Value == null ||
            selectionResult.Value.Count == 0)
        {
          return;
        }

        PromptPointResult basePointResult = ed.GetPoint(
          $"\nSpecify base point for {degrees:0}-degree rotation: ");

        if (basePointResult.Status != PromptStatus.OK)
        {
          return;
        }

        Matrix3d ucsToWcs = ed.CurrentUserCoordinateSystem;
        Point3d basePoint = basePointResult.Value.TransformBy(ucsToWcs);
        Vector3d rotationAxis = Vector3d.ZAxis.TransformBy(ucsToWcs).GetNormal();
        double angle = degrees * Math.PI / 180.0;
        Matrix3d rotation = Matrix3d.Rotation(angle, rotationAxis, basePoint);

        using (Transaction transaction = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId objectId in selectionResult.Value.GetObjectIds())
          {
            Entity entity = transaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
            entity?.TransformBy(rotation);
          }

          transaction.Commit();
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to rotate the selected objects: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(Array.Empty<ObjectId>());
      }
    }
  }
}
