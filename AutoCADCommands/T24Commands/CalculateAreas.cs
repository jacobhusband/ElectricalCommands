using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using ElectricalCommands;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("AREALABEL", CommandFlags.UsePickSet)]
    [CommandMethod("QA", CommandFlags.UsePickSet)]
    public void AREACALCULATOR()
    {
      var (doc, db, ed) = Globals.GetGlobals();

      try
      {
        SelectionSet sset;
        PromptSelectionResult selRes = ed.SelectImplied();
        if (selRes.Status == PromptStatus.OK)
        {
          // Use the PickFirst selection
          sset = selRes.Value;
        }
        else
        {
          // If no PickFirst selection, prompt for selection
          PromptSelectionOptions opts = new PromptSelectionOptions();
          opts.MessageForAdding = "Select polylines or rectangles: ";
          opts.AllowDuplicates = false;
          opts.RejectObjectsOnLockedLayers = true;
          TypedValue[] filterList = new TypedValue[]
          {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
          };
          SelectionFilter filter = new SelectionFilter(filterList);
          selRes = ed.GetSelection(opts, filter);
          if (selRes.Status != PromptStatus.OK)
            return;
          sset = selRes.Value;
        }

        using (Transaction tr = doc.TransactionManager.StartTransaction())
        {
          int processedCount = 0;
          foreach (ObjectId objId in sset.GetObjectIds())
          {
            var obj = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
            if (obj == null)
            {
              ed.WriteMessage("\nSelected object is not a valid entity.");
              continue;
            }

            Autodesk.AutoCAD.DatabaseServices.Polyline polyline =
                obj as Autodesk.AutoCAD.DatabaseServices.Polyline;
            if (polyline != null)
            {
              double area = polyline.Area;
              area /= 144; // Converting from square inches to square feet
              ed.WriteMessage(
                  $"\nThe area of the selected polyline is: {area:F2} sq ft"
              );

              // Get the bounding box of the polyline
              Extents3d bounds = (Extents3d)polyline.Bounds;
              // Calculate the center of the bounding box
              Point3d center = new Point3d(
                  (bounds.MinPoint.X + bounds.MaxPoint.X) / 2,
                  (bounds.MinPoint.Y + bounds.MaxPoint.Y) / 2,
                  0
              );

              // Check if the center of the bounding box lies within the polyline. If not, use the first vertex.
              if (!IsPointInside(polyline, center))
              {
                center = polyline.GetPoint3dAt(0);
              }

              DBText text = new DBText
              {
                Height = 9,
                TextString = $"{Math.Ceiling(area)} sq ft",
                Rotation = 0,
                HorizontalMode = TextHorizontalMode.TextCenter,
                VerticalMode = TextVerticalMode.TextVerticalMid,
                Layer = "0"
              };
              text.Position = center;
              text.AlignmentPoint = center;

              var currentSpace = (BlockTableRecord)
                  tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
              currentSpace.AppendEntity(text);
              tr.AddNewlyCreatedDBObject(text, true);

              processedCount++;
            }
            else
            {
              ed.WriteMessage("\nSelected object is not a polyline.");
              continue;
            }
          }
          tr.Commit();
          ed.WriteMessage($"\nAREACALCULATOR command completed successfully. Processed {processedCount} polyline(s).");
        }

        // Clear the PickFirst selection set
        ed.SetImpliedSelection(new ObjectId[0]);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nError: {ex.Message}");
      }
    }

    public static bool IsPointInside(Polyline polyline, Point3d point)
    {
      int numIntersections = 0;
      for (int i = 0; i < polyline.NumberOfVertices; i++)
      {
        Point3d point1 = polyline.GetPoint3dAt(i);
        Point3d point2 = polyline.GetPoint3dAt((i + 1) % polyline.NumberOfVertices); // Get next point, or first point if we're at the end

        // Check if point is on an horizontal segment
        if (
            point1.Y == point2.Y
            && point1.Y == point.Y
            && point.X > Math.Min(point1.X, point2.X)
            && point.X < Math.Max(point1.X, point2.X)
        )
        {
          return true;
        }

        if (
            point.Y > Math.Min(point1.Y, point2.Y)
            && point.Y <= Math.Max(point1.Y, point2.Y)
            && point.X <= Math.Max(point1.X, point2.X)
            && point1.Y != point2.Y
        )
        {
          double xinters =
              (point.Y - point1.Y) * (point2.X - point1.X) / (point2.Y - point1.Y)
              + point1.X;

          // Check if point is on the polygon boundary (other than horizontal)
          if (Math.Abs(point.X - xinters) < Double.Epsilon)
          {
            return true;
          }

          // Count intersections
          if (point.X < xinters)
          {
            numIntersections++;
          }
        }
      }
      // If the number of intersections is odd, the point is inside.
      return numIntersections % 2 != 0;
    }
  }
}
