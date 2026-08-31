using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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

        PromptPointResult basePointResult = ed.GetPoint(
          new PromptPointOptions(
            "\nSpecify the common base point for the saved room boundaries: "));
        if (basePointResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nAREALABEL canceled before room data was saved.");
          return;
        }

        Point3d basePoint = basePointResult.Value.TransformBy(
          ed.CurrentUserCoordinateSystem);
        List<ElectricalDrawingSettingsStore.RoomBoundarySetting>
          savedRooms =
            new List<ElectricalDrawingSettingsStore.RoomBoundarySetting>();
        int unnamedRoomCount = 0;

        using (Transaction tr = doc.TransactionManager.StartTransaction())
        {
          List<AreaLabelTextCandidate> roomTextCandidates =
            ReadAreaLabelTextCandidates(db, tr);
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

              if (polyline.Closed && polyline.NumberOfVertices >= 3)
              {
                string roomName = ResolveAreaLabelRoomName(
                  polyline,
                  center,
                  roomTextCandidates);
                if (roomName.Length == 0)
                {
                  unnamedRoomCount++;
                }

                savedRooms.Add(
                  new ElectricalDrawingSettingsStore.RoomBoundarySetting
                  {
                    Name = roomName,
                    SourceHandle = polyline.Handle.ToString(),
                    RelativeBoundary = BuildRelativeRoomBoundary(
                      polyline,
                      basePoint),
                  });
              }
              else
              {
                ed.WriteMessage(
                  $"\nPolyline {polyline.Handle} is not closed and was not " +
                  "saved as a room boundary.");
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
          ed.WriteMessage(
            $"\nAREALABEL processed {processedCount} polyline(s).");
        }

        if (savedRooms.Count > 0)
        {
          ElectricalDrawingSettingsStore.WriteRoomBoundaries(
            db,
            basePoint,
            savedRooms);
          ed.WriteMessage(
            $"\nSaved {savedRooms.Count} room boundary(ies) relative to " +
            $"base point {FormatAreaLabelPoint(basePoint)} for RC." +
            (unnamedRoomCount > 0
              ? $" {unnamedRoomCount} boundary(ies) did not contain " +
                "recognizable room-name text and must be corrected before " +
                "RC can use receptacles in those rooms."
              : string.Empty));
        }
        else
        {
          ed.WriteMessage(
            "\nNo closed room boundaries were available to save for RC.");
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

    private static List<AreaLabelTextCandidate> ReadAreaLabelTextCandidates(
      Database database,
      Transaction transaction)
    {
      List<AreaLabelTextCandidate> candidates =
        new List<AreaLabelTextCandidate>();
      BlockTableRecord currentSpace = transaction.GetObject(
        database.CurrentSpaceId,
        OpenMode.ForRead) as BlockTableRecord;
      if (currentSpace == null)
      {
        return candidates;
      }

      foreach (ObjectId objectId in currentSpace)
      {
        Entity entity = transaction.GetObject(
          objectId,
          OpenMode.ForRead,
          false) as Entity;
        string text = string.Empty;
        Point3d position = Point3d.Origin;
        if (entity is DBText dbText)
        {
          text = dbText.TextString;
          position = dbText.Position;
        }
        else if (entity is MText mText)
        {
          text = mText.Text;
          position = mText.Location;
        }

        text = NormalizeAreaLabelRoomText(text);
        if (text.Length == 0 || IsAreaMeasurementText(text))
        {
          continue;
        }

        candidates.Add(new AreaLabelTextCandidate
        {
          Text = text,
          Position = position,
        });
      }
      return candidates;
    }

    private static string ResolveAreaLabelRoomName(
      Polyline polyline,
      Point3d center,
      List<AreaLabelTextCandidate> candidates)
    {
      List<AreaLabelTextCandidate> inside = candidates
        .Where(candidate => IsPointInside(polyline, candidate.Position))
        .OrderBy(candidate => candidate.Position.DistanceTo(center))
        .ToList();
      if (inside.Count == 0)
      {
        return string.Empty;
      }

      AreaLabelTextCandidate descriptiveText = inside.FirstOrDefault(
        candidate =>
          Regex.IsMatch(candidate.Text, "[A-Z]", RegexOptions.IgnoreCase) &&
          !IsLikelyRoomNumber(candidate.Text));
      if (descriptiveText == null)
      {
        return inside[0].Text;
      }

      string roomName = descriptiveText.Text;
      if (!Regex.IsMatch(roomName, @"\d"))
      {
        AreaLabelTextCandidate numberText = inside.FirstOrDefault(
          candidate => IsLikelyRoomNumber(candidate.Text));
        if (numberText != null)
        {
          roomName += " " + numberText.Text;
        }
      }
      return NormalizeAreaLabelRoomText(roomName);
    }

    private static bool IsLikelyRoomNumber(string text)
    {
      return Regex.IsMatch(
        text ?? string.Empty,
        @"^[A-Z]?\d{1,5}[A-Z]?$",
        RegexOptions.IgnoreCase);
    }

    private static bool IsAreaMeasurementText(string text)
    {
      return Regex.IsMatch(
        text ?? string.Empty,
        @"^\s*[\d,.]+\s*(?:SQ\.?\s*FT\.?|SF|SQUARE\s+FEET)\s*$",
        RegexOptions.IgnoreCase);
    }

    private static string NormalizeAreaLabelRoomText(string text)
    {
      return Regex.Replace(
        text ?? string.Empty,
        @"\s+",
        " ").Trim().ToUpperInvariant();
    }

    private static List<Point2d> BuildRelativeRoomBoundary(
      Polyline polyline,
      Point3d basePoint)
    {
      List<Point2d> points = new List<Point2d>();
      int segmentCount = polyline.Closed
        ? polyline.NumberOfVertices
        : Math.Max(0, polyline.NumberOfVertices - 1);
      for (int segmentIndex = 0;
           segmentIndex < segmentCount;
           segmentIndex++)
      {
        double includedAngle = Math.Abs(
          4.0 * Math.Atan(polyline.GetBulgeAt(segmentIndex)));
        int subdivisions = Math.Max(
          1,
          (int)Math.Ceiling(includedAngle / (Math.PI / 18.0)));
        for (int subdivision = 0;
             subdivision < subdivisions;
             subdivision++)
        {
          double parameter =
            segmentIndex + subdivision / (double)subdivisions;
          Point3d point = polyline.GetPointAtParameter(parameter);
          points.Add(new Point2d(
            point.X - basePoint.X,
            point.Y - basePoint.Y));
        }
      }
      return points;
    }

    private static string FormatAreaLabelPoint(Point3d point)
    {
      return $"({point.X:0.###}, {point.Y:0.###}, {point.Z:0.###})";
    }

    private sealed class AreaLabelTextCandidate
    {
      internal string Text { get; set; } = string.Empty;
      internal Point3d Position { get; set; }
    }
  }
}
