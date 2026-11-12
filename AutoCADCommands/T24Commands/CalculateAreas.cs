using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
        List<AreaLabelRoomItem> roomItems = LoadSavedAreaLabelRooms(
          db,
          ed,
          out Point3d? existingBasePoint);

        HashSet<ObjectId> existingIds = new HashSet<ObjectId>(
          roomItems.Select(r => r.ObjectId));
        HashSet<string> existingHandles = new HashSet<string>(
          roomItems.Select(r => r.SourceHandle),
          StringComparer.OrdinalIgnoreCase);

        SelectionSet selectionSet = null;
        PromptSelectionResult selectionResult = ed.SelectImplied();
        if (selectionResult.Status == PromptStatus.OK &&
            selectionResult.Value != null &&
            selectionResult.Value.Count > 0)
        {
          selectionSet = selectionResult.Value;
        }
        else if (roomItems.Count == 0)
        {
          // Rooms already saved go straight to the review window, where
          // more polylines can be added and the base point changed.
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
          if (selectionResult.Status == PromptStatus.OK &&
              selectionResult.Value != null &&
              selectionResult.Value.Count > 0)
          {
            selectionSet = selectionResult.Value;
          }
          else
          {
            return;
          }
        }

        if (selectionSet != null)
        {
          using (Transaction transaction =
            db.TransactionManager.StartOpenCloseTransaction())
          {
            foreach (ObjectId objectId in selectionSet.GetObjectIds())
            {
              if (objectId.IsNull || !objectId.IsValid || objectId.IsErased)
              {
                continue;
              }

              if (existingIds.Contains(objectId))
              {
                continue;
              }

              Polyline polyline = transaction.GetObject(
                objectId,
                OpenMode.ForRead,
                false) as Polyline;
              if (polyline == null)
              {
                ed.WriteMessage(
                  $"\nIgnored selected object {objectId.Handle}; it is not a " +
                  "lightweight polyline.");
                continue;
              }
              if (polyline.NumberOfVertices < 3)
              {
                ed.WriteMessage(
                  $"\nIgnored polyline {polyline.Handle}; room boundaries must " +
                  "contain at least three vertices.");
                continue;
              }

              string handleStr = polyline.Handle.ToString();
              if (existingHandles.Contains(handleStr))
              {
                continue;
              }

              double squareFeet = polyline.Area / 144.0;
              int roomNumber = roomItems.Count + 1;
              string defaultName = $"Polyline {roomNumber}";
              string roomName = defaultName;
              string group = string.Empty;

              if (RoomBoundaryMetadataStore.TryRead(
                polyline,
                transaction,
                out var existingMetadata) &&
                !string.IsNullOrWhiteSpace(existingMetadata.Name))
              {
                roomName = existingMetadata.Name;
                group = existingMetadata.Group;
              }

              roomItems.Add(new AreaLabelRoomItem
              {
                ObjectId = objectId,
                SourceLabel = defaultName,
                SourceHandle = handleStr,
                DefaultRoomName = defaultName,
                RoomName = roomName,
                SquareFeet = squareFeet,
                PerimeterFeet = CalculateAreaLabelPerimeterFeet(polyline),
                Group = group,
              });
              existingIds.Add(objectId);
              existingHandles.Add(handleStr);
            }
          }
        }

        if (roomItems.Count == 0)
        {
          ed.WriteMessage(
            "\nAREALABEL requires at least one polyline with three " +
            "or more vertices.");
          return;
        }

        // Remove PickFirst grips so the naming window can show clear,
        // individually highlighted boundaries.
        ed.SetImpliedSelection(new ObjectId[0]);
        List<ObjectId> highlightedBoundaryIds = new List<ObjectId>();
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
          ed,
          db,
          existingBasePoint);

        namingWindow.SelectedBoundariesChanged += selectedIds =>
        {
          foreach (ObjectId id in highlightedBoundaryIds)
          {
            SetAreaLabelBoundaryHighlight(db, id, false);
          }

          highlightedBoundaryIds = selectedIds?
            .Where(id => !id.IsNull && id.IsValid && !id.IsErased)
            .ToList() ?? new List<ObjectId>();

          foreach (ObjectId id in highlightedBoundaryIds)
          {
            SetAreaLabelBoundaryHighlight(db, id, true);
          }
        };

        bool? accepted;
        try
        {
          accepted = AcApplication.ShowModalWindow(namingWindow);
        }
        finally
        {
          foreach (ObjectId id in highlightedBoundaryIds)
          {
            SetAreaLabelBoundaryHighlight(db, id, false);
          }
        }

        if (accepted != true)
        {
          ed.WriteMessage("\nAREALABEL room naming canceled.");
          return;
        }

        // The saved base point (or one picked in the window) is reused; the
        // prompt only appears when the drawing has never had one.
        Point3d basePoint;
        if (namingWindow.BasePoint.HasValue)
        {
          basePoint = namingWindow.BasePoint.Value;
        }
        else
        {
          PromptPointResult basePointResult = ed.GetPoint(
            new PromptPointOptions(
              "\nSpecify the common base point for the saved room boundaries: "));
          if (basePointResult.Status != PromptStatus.OK)
          {
            ed.WriteMessage("\nAREALABEL canceled before room data was saved.");
            return;
          }
          basePoint = basePointResult.Value.TransformBy(
            ed.CurrentUserCoordinateSystem);
        }

        List<ElectricalDrawingSettingsStore.RoomBoundarySetting> savedRooms =
          new List<ElectricalDrawingSettingsStore.RoomBoundarySetting>();

        using (Transaction transaction =
          doc.TransactionManager.StartTransaction())
        {
          // Clean up metadata from polylines that were explicitly removed in the UI
          foreach (ObjectId removedId in namingWindow.RemovedObjectIds)
          {
            if (removedId.IsNull || !removedId.IsValid || removedId.IsErased)
            {
              continue;
            }

            Polyline removedPolyline = transaction.GetObject(
              removedId,
              OpenMode.ForWrite,
              false) as Polyline;
            if (removedPolyline != null)
            {
              RoomBoundaryMetadataStore.Remove(removedPolyline, transaction);
            }
          }

          // Save metadata and settings for all active rooms
          foreach (AreaLabelRoomItem roomItem in namingWindow.Rooms)
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
                Group = roomItem.Group,
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
        WriteAreaLabelGroupTotals(ed, namingWindow.Rooms);
        ed.SetImpliedSelection(new ObjectId[0]);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nAREALABEL failed: {ex.Message}");
      }
    }

    private static List<AreaLabelRoomItem> LoadSavedAreaLabelRooms(
      Database database,
      Editor editor,
      out Point3d? existingBasePoint)
    {
      existingBasePoint = null;
      List<AreaLabelRoomItem> rooms = new List<AreaLabelRoomItem>();
      if (!ElectricalDrawingSettingsStore.TryReadRoomBoundaries(
        database,
        out var savedRoomData) ||
        savedRoomData == null ||
        savedRoomData.Rooms == null ||
        savedRoomData.Rooms.Count == 0)
      {
        return rooms;
      }

      existingBasePoint = savedRoomData.BasePoint;
      HashSet<string> seenHandles = new HashSet<string>(
        StringComparer.OrdinalIgnoreCase);

      using (Transaction transaction =
        database.TransactionManager.StartOpenCloseTransaction())
      {
        foreach (var roomSetting in savedRoomData.Rooms)
        {
          if (string.IsNullOrWhiteSpace(roomSetting.SourceHandle))
          {
            continue;
          }

          if (seenHandles.Contains(roomSetting.SourceHandle))
          {
            continue;
          }

          ObjectId polylineId = ResolvePolylineByHandle(
            database,
            roomSetting.SourceHandle);
          if (polylineId.IsNull || !polylineId.IsValid || polylineId.IsErased)
          {
            continue;
          }

          Polyline polyline = transaction.GetObject(
            polylineId,
            OpenMode.ForRead,
            false) as Polyline;
          if (polyline == null || polyline.IsErased || polyline.NumberOfVertices < 3)
          {
            continue;
          }

          double squareFeet = polyline.Area / 144.0;
          string roomName = roomSetting.Name ?? string.Empty;
          string group = string.Empty;
          if (RoomBoundaryMetadataStore.TryRead(
            polyline,
            transaction,
            out var metadata) &&
            !string.IsNullOrWhiteSpace(metadata.Name))
          {
            roomName = metadata.Name;
            group = metadata.Group;
          }

          int roomNumber = rooms.Count + 1;
          string defaultName = $"Polyline {roomNumber}";
          if (string.IsNullOrWhiteSpace(roomName))
          {
            roomName = defaultName;
          }

          rooms.Add(new AreaLabelRoomItem
          {
            ObjectId = polylineId,
            SourceLabel = defaultName,
            SourceHandle = polyline.Handle.ToString(),
            DefaultRoomName = defaultName,
            RoomName = roomName,
            SquareFeet = squareFeet,
            PerimeterFeet = CalculateAreaLabelPerimeterFeet(polyline),
            Group = group,
          });
          seenHandles.Add(roomSetting.SourceHandle);
        }
      }

      return rooms;
    }

    private static ObjectId ResolvePolylineByHandle(
      Database database,
      string handleStr)
    {
      if (database == null || string.IsNullOrWhiteSpace(handleStr))
      {
        return ObjectId.Null;
      }

      if (!long.TryParse(
        handleStr,
        NumberStyles.HexNumber,
        CultureInfo.InvariantCulture,
        out long handleVal))
      {
        return ObjectId.Null;
      }

      try
      {
        return database.GetObjectId(false, new Handle(handleVal), 0);
      }
      catch
      {
        return ObjectId.Null;
      }
    }

    /// <summary>
    /// Perimeter in feet. AutoCAD reports an open polyline's area as if it
    /// were closed, so the closing segment is included here as well to keep
    /// the perimeter consistent with the area.
    /// </summary>
    internal static double CalculateAreaLabelPerimeterFeet(Polyline polyline)
    {
      double lengthInches = polyline.Length;
      if (!polyline.Closed && polyline.NumberOfVertices > 1)
      {
        lengthInches += polyline.GetPoint2dAt(polyline.NumberOfVertices - 1)
          .GetDistanceTo(polyline.GetPoint2dAt(0));
      }
      return lengthInches / 12.0;
    }

    private static void WriteAreaLabelGroupTotals(
      Editor editor,
      IEnumerable<AreaLabelRoomItem> rooms)
    {
      List<AreaLabelRoomItem> roomList = rooms.ToList();
      if (!roomList.Any(r => !string.IsNullOrWhiteSpace(r.Group)))
      {
        return;
      }

      List<AreaLabelGroupTotal> totals = AreaLabelGroupTotal.Build(roomList);

      foreach (AreaLabelGroupTotal total in totals)
      {
        editor.WriteMessage(
          $"\n  {total.GroupName}: {total.Count} polyline(s), " +
          $"{total.AreaText}, {total.PerimeterText} perimeter");
      }
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

    internal static string FormatAreaLabelPoint(Point3d point)
    {
      return $"({point.X:0.###}, {point.Y:0.###}, {point.Z:0.###})";
    }
  }
}
