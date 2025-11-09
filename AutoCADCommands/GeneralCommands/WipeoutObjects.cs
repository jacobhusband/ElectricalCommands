using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("WIPEOUTOBJECTS", CommandFlags.UsePickSet)]
    public static void WO()
    {
      Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      // First, try to get the implied selection (PickFirst)
      PromptSelectionResult selectionResult = ed.SelectImplied();

      // If no implied selection, prompt the user to select objects
      if (selectionResult.Status != PromptStatus.OK)
      {
        PromptSelectionOptions pso = new PromptSelectionOptions();
        pso.MessageForAdding = "Select objects for WO command: ";
        selectionResult = ed.GetSelection(pso);
      }

      if (selectionResult.Status == PromptStatus.OK)
      {
        SelectionSet selectionSet = selectionResult.Value;
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
          BlockTableRecord currentSpace = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

          foreach (SelectedObject selObj in selectionSet)
          {
            Entity ent = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
            if (ent != null)
            {
              Wipeout wo = null;

              if (ent is DBText dbText)
              {
                wo = CreateWipeoutForText(dbText, 0.25);
              }
              else if (ent is MText mText)
              {
                wo = CreateWipeoutForText(mText, 0.25);
              }
              else if (ent is Autodesk.AutoCAD.DatabaseServices.Table table)
              {
                wo = CreateWipeoutForTable(table);
              }
              else if (ent is Polyline pline)
              {
                wo = CreateWipeoutForPolyline(pline);
              }

              if (wo != null)
              {
                currentSpace.AppendEntity(wo);
                trans.AddNewlyCreatedDBObject(wo, true);

                // Send the original object to the front
                DrawOrderTable dot = (DrawOrderTable)trans.GetObject(currentSpace.DrawOrderTableId, OpenMode.ForWrite);
                ObjectIdCollection ids = new ObjectIdCollection { selObj.ObjectId };
                dot.MoveToTop(ids);
              }
            }
          }

          trans.Commit();
        }
      }
      else
      {
        ed.WriteMessage("\nNo objects were selected.");
      }

      // Clear the PickFirst selection set
      ed.SetImpliedSelection(new ObjectId[0]);
    }

    private static Wipeout CreateWipeoutForText(Entity textEntity, double paddingFactor)
    {
      Extents3d extents = textEntity.GeometricExtents;
      double height = (textEntity is DBText dbText) ? dbText.Height : ((MText)textEntity).TextHeight;
      double padding = height * paddingFactor;

      Point2d minPt = new Point2d(extents.MinPoint.X - padding, extents.MinPoint.Y - padding);
      Point2d maxPt = new Point2d(extents.MaxPoint.X + padding, extents.MaxPoint.Y + padding);

      return CreateWipeoutFromPoints(minPt, maxPt);
    }

    private static Wipeout CreateWipeoutForTable(Autodesk.AutoCAD.DatabaseServices.Table table)
    {
      Extents3d extents = table.GeometricExtents;
      return CreateWipeoutFromPoints(new Point2d(extents.MinPoint.X, extents.MinPoint.Y),
                                     new Point2d(extents.MaxPoint.X, extents.MaxPoint.Y));
    }

    private static Wipeout CreateWipeoutForPolyline(Polyline pline)
    {
      Extents3d extents = pline.GeometricExtents;
      return CreateWipeoutFromPoints(new Point2d(extents.MinPoint.X, extents.MinPoint.Y),
                                     new Point2d(extents.MaxPoint.X, extents.MaxPoint.Y));
    }

    private static Wipeout CreateWipeoutFromPoints(Point2d minPt, Point2d maxPt)
    {
      Point2dCollection pts = new Point2dCollection
      {
        minPt,
        new Point2d(maxPt.X, minPt.Y),
        maxPt,
        new Point2d(minPt.X, maxPt.Y),
        minPt // Close the loop
      };

      Wipeout wo = new Wipeout();
      wo.SetDatabaseDefaults();
      wo.SetFrom(pts, Vector3d.ZAxis);

      return wo;
    }


  }
}
