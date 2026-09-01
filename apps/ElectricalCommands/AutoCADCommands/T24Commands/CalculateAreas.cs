using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Application;

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
        SelectionSet selectionSet;
        PromptSelectionResult selectionResult = ed.SelectImplied();
        if (selectionResult.Status == PromptStatus.OK &&
            selectionResult.Value != null)
        {
          selectionSet = selectionResult.Value;
        }
        else
        {
          PromptSelectionOptions options = new PromptSelectionOptions
          {
            MessageForAdding = "\nSelect room polylines: ",
            AllowDuplicates = false,
            RejectObjectsOnLockedLayers = true,
          };
          SelectionFilter filter = new SelectionFilter(
            new[]
            {
              new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
            });
          selectionResult = ed.GetSelection(options, filter);
          if (selectionResult.Status != PromptStatus.OK ||
              selectionResult.Value == null)
          {
            return;
          }
          selectionSet = selectionResult.Value;
        }

        List<AreaLabelRoomItem> roomItems = BuildAreaLabelRoomItems(
          db,
          ed,
          selectionSet);
        if (roomItems.Count == 0)
        {
          ed.WriteMessage(
            "\nAREALABEL requires at least one polyline with three " +
            "or more vertices.");
          return;
        }

        // Remove PickFirst grips so the naming window can show one clear,
        // individually highlighted boundary at a time.
        ed.SetImpliedSelection(new ObjectId[0]);
        ObjectId highlightedBoundaryId = ObjectId.Null;
        string drawingDirectory = null;
        try
        {
          if (!string.IsNullOrWhiteSpace(db.Filename))
          {
            drawingDirectory = System.IO.Path.GetDirectoryName(db.Filename);
          }
        }
        catch
        {
          // Ignore drawing path resolution exceptions
        }
        RoomNamingWindow namingWindow = new RoomNamingWindow(
          roomItems,
          drawingDirectory,
          ed);
        namingWindow.SelectedBoundaryChanged += selectedId =>
        {
          if (selectedId == highlightedBoundaryId)
          {
            return;
          }
          SetAreaLabelBoundaryHighlight(
            db,
            highlightedBoundaryId,
            false);
          highlightedBoundaryId = selectedId;
          SetAreaLabelBoundaryHighlight(db, highlightedBoundaryId, true);
        };

        bool? accepted;
        try
        {
          accepted = AcApplication.ShowModalWindow(namingWindow);
        }
        finally
        {
          SetAreaLabelBoundaryHighlight(
            db,
            highlightedBoundaryId,
            false);
        }
        if (accepted != true)
        {
          ed.WriteMessage("\nAREALABEL room naming canceled.");
          return;
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
        List<ElectricalDrawingSettingsStore.RoomBoundarySetting> savedRooms =
          new List<ElectricalDrawingSettingsStore.RoomBoundarySetting>();

        using (Transaction transaction =
          doc.TransactionManager.StartTransaction())
        {
          foreach (AreaLabelRoomItem roomItem in roomItems)
          {
            Polyline polyline = transaction.GetObject(
              roomItem.ObjectId,
              OpenMode.ForWrite,
              false) as Polyline;
            if (polyline == null ||
                polyline.NumberOfVertices < 3)
            {
              throw new InvalidOperationException(
                $"Room boundary {roomItem.SourceHandle} is no longer a " +
                "valid polyline with at least three vertices.");
            }

            roomItem.SquareFeet = polyline.Area / 144.0;
            Point3d labelLocation = ResolveAreaLabelLocation(polyline);
            Point2d relativeLocation = new Point2d(
              labelLocation.X - basePoint.X,
              labelLocation.Y - basePoint.Y);
            RoomBoundaryMetadataStore.Write(
              polyline,
              transaction,
              new RoomBoundaryMetadataStore.RoomMetadata
              {
                Name = roomItem.RoomName,
                SquareFeet = roomItem.SquareFeet,
                BasePoint = basePoint,
                RelativeLocation = relativeLocation,
              });

            savedRooms.Add(
              new ElectricalDrawingSettingsStore.RoomBoundarySetting
              {
                Name = roomItem.RoomName,
                SourceHandle = polyline.Handle.ToString(),
                SquareFeet = roomItem.SquareFeet,
                RelativeLocation = relativeLocation,
                RelativeBoundary = BuildRelativeRoomBoundary(
                  polyline,
                  basePoint),
              });
          }
          transaction.Commit();
        }

        ElectricalDrawingSettingsStore.WriteRoomBoundaries(
          db,
          basePoint,
          savedRooms);
        ed.WriteMessage(
          $"\nAREALABEL saved invisible metadata on {savedRooms.Count} " +
          "room polyline(s): names, current square footage, and locations " +
          $"relative to base point {FormatAreaLabelPoint(basePoint)}. No " +
          "drawing text was created.");
        ed.SetImpliedSelection(new ObjectId[0]);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nAREALABEL failed: {ex.Message}");
      }
    }

    private static List<AreaLabelRoomItem> BuildAreaLabelRoomItems(
      Database database,
      Editor editor,
      SelectionSet selectionSet)
    {
      List<AreaLabelRoomItem> rooms = new List<AreaLabelRoomItem>();
      Dictionary<string, string> legacyRoomNames =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
      if (ElectricalDrawingSettingsStore.TryReadRoomBoundaries(
        database,
        out var legacyRoomData))
      {
        foreach (ElectricalDrawingSettingsStore.RoomBoundarySetting legacyRoom
          in legacyRoomData.Rooms)
        {
          if (!string.IsNullOrWhiteSpace(legacyRoom.SourceHandle) &&
              !string.IsNullOrWhiteSpace(legacyRoom.Name))
          {
            legacyRoomNames[legacyRoom.SourceHandle] = legacyRoom.Name;
          }
        }
      }

      using (Transaction transaction =
        database.TransactionManager.StartOpenCloseTransaction())
      {
        foreach (ObjectId objectId in selectionSet.GetObjectIds())
        {
          Polyline polyline = transaction.GetObject(
            objectId,
            OpenMode.ForRead,
            false) as Polyline;
          if (polyline == null)
          {
            editor.WriteMessage(
              $"\nIgnored selected object {objectId.Handle}; it is not a " +
              "lightweight polyline.");
            continue;
          }
          if (polyline.NumberOfVertices < 3)
          {
            editor.WriteMessage(
              $"\nIgnored polyline {polyline.Handle}; room boundaries must " +
              "contain at least three vertices.");
            continue;
          }

          double squareFeet = polyline.Area / 144.0;
          int roomNumber = rooms.Count + 1;
          string defaultName = $"Polyline {roomNumber}";
          string roomName = defaultName;
          if (RoomBoundaryMetadataStore.TryRead(
            polyline,
            transaction,
            out var existingMetadata))
          {
            roomName = existingMetadata.Name;
          }
          else if (legacyRoomNames.TryGetValue(
            polyline.Handle.ToString(),
            out string legacyRoomName))
          {
            roomName = legacyRoomName;
          }
          rooms.Add(new AreaLabelRoomItem
          {
            ObjectId = objectId,
            SourceLabel = defaultName,
            SourceHandle = polyline.Handle.ToString(),
            DefaultRoomName = defaultName,
            RoomName = roomName,
            SquareFeet = squareFeet,
          });
        }
      }
      return rooms;
    }

    private static Point3d ResolveAreaLabelLocation(Polyline polyline)
    {
      Extents3d bounds = polyline.GeometricExtents;
      Point3d center = new Point3d(
        (bounds.MinPoint.X + bounds.MaxPoint.X) / 2.0,
        (bounds.MinPoint.Y + bounds.MaxPoint.Y) / 2.0,
        (bounds.MinPoint.Z + bounds.MaxPoint.Z) / 2.0);
      return IsPointInside(polyline, center)
        ? center
        : polyline.GetPoint3dAt(0);
    }

    private static void SetAreaLabelBoundaryHighlight(
      Database database,
      ObjectId objectId,
      bool highlighted)
    {
      if (objectId.IsNull || !objectId.IsValid || objectId.IsErased)
      {
        return;
      }

      try
      {
        using (Transaction transaction =
          database.TransactionManager.StartOpenCloseTransaction())
        {
          Entity entity = transaction.GetObject(
            objectId,
            OpenMode.ForRead,
            false) as Entity;
          if (entity != null)
          {
            if (highlighted)
            {
              entity.Highlight();
            }
            else
            {
              entity.Unhighlight();
            }
          }
        }
        AcApplication.UpdateScreen();
      }
      catch
      {
        // A stale highlight must not interrupt room naming.
      }
    }

    public static bool IsPointInside(Polyline polyline, Point3d point)
    {
      int numIntersections = 0;
      for (int index = 0; index < polyline.NumberOfVertices; index++)
      {
        Point3d first = polyline.GetPoint3dAt(index);
        Point3d second = polyline.GetPoint3dAt(
          (index + 1) % polyline.NumberOfVertices);

        if (first.Y == second.Y &&
            first.Y == point.Y &&
            point.X > Math.Min(first.X, second.X) &&
            point.X < Math.Max(first.X, second.X))
        {
          return true;
        }

        if (point.Y > Math.Min(first.Y, second.Y) &&
            point.Y <= Math.Max(first.Y, second.Y) &&
            point.X <= Math.Max(first.X, second.X) &&
            first.Y != second.Y)
        {
          double intersectionX =
            (point.Y - first.Y) * (second.X - first.X) /
            (second.Y - first.Y) + first.X;
          if (Math.Abs(point.X - intersectionX) < double.Epsilon)
          {
            return true;
          }
          if (point.X < intersectionX)
          {
            numIntersections++;
          }
        }
      }
      return numIntersections % 2 != 0;
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
      if (!polyline.Closed && polyline.NumberOfVertices > 0)
      {
        Point3d finalPoint = polyline.GetPoint3dAt(
          polyline.NumberOfVertices - 1);
        points.Add(new Point2d(
          finalPoint.X - basePoint.X,
          finalPoint.Y - basePoint.Y));
      }
      return points;
    }

    private static string FormatAreaLabelPoint(Point3d point)
    {
      return $"({point.X:0.###}, {point.Y:0.###}, {point.Z:0.###})";
    }
  }
}
