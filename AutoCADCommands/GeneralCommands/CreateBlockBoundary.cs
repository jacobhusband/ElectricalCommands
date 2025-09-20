using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const double VertexTolerance = 1e-6;

    [CommandMethod("BLOCKBOUNDARY", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public void CreateBoundaryAroundBlockReferences()
    {
      var doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      var db = doc.Database;
      var ed = doc.Editor;

      SelectionSet selection = null;

      var implied = ed.SelectImplied();
      if (implied.Status == PromptStatus.OK && implied.Value != null && implied.Value.Count > 0)
      {
        selection = implied.Value;
        ed.SetImpliedSelection(Array.Empty<ObjectId>());
      }

      if (selection == null)
      {
        var selOpts = new PromptSelectionOptions
        {
          MessageForAdding = "\nSelect block references in model space: ",
          MessageForRemoval = "\nRemove block references: ",
          AllowDuplicates = false,
          RejectObjectsOnLockedLayers = true
        };

        var filter = new SelectionFilter(new[]
        {
          new TypedValue((int)DxfCode.Start, "INSERT")
        });

        var selRes = ed.GetSelection(selOpts, filter);
        if (selRes.Status != PromptStatus.OK)
        {
          return;
        }

        selection = selRes.Value;
      }

      if (selection == null || selection.Count == 0)
      {
        ed.WriteMessage("\nNo block references selected.");
        return;
      }

      var paddingOptions = new PromptDoubleOptions("\nEnter padding distance: ")
      {
        AllowNegative = false,
        AllowZero = true,
        DefaultValue = 0.0,
        UseDefaultValue = true
      };

      var paddingResult = ed.GetDouble(paddingOptions);
      if (paddingResult.Status != PromptStatus.OK)
      {
        return;
      }

      double padding = paddingResult.Value;

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var modelId = bt[BlockTableRecord.ModelSpace];
        var model = (BlockTableRecord)tr.GetObject(modelId, OpenMode.ForRead);

        var points = new List<Point2d>();
        var comparer = new Point2dEqualityComparer(VertexTolerance);
        var dedupe = new HashSet<Point2d>(comparer);

        int included = 0;
        int skipped = 0;
        double elevation = 0.0;
        bool first = true;

        foreach (var id in selection.GetObjectIds())
        {
          if (id == ObjectId.Null) continue;

          var obj = tr.GetObject(id, OpenMode.ForRead, false);
          if (obj is not BlockReference blockRef)
          {
            skipped++;
            continue;
          }

          if (blockRef.OwnerId != modelId)
          {
            skipped++;
            continue;
          }

          var extents = TryGetEntityExtents(blockRef);
          if (extents == null)
          {
            skipped++;
            continue;
          }

          var ext = extents.Value;
          double minX = ext.MinPoint.X - padding;
          double minY = ext.MinPoint.Y - padding;
          double maxX = ext.MaxPoint.X + padding;
          double maxY = ext.MaxPoint.Y + padding;

          AddIfUnique(points, dedupe, new Point2d(minX, minY));
          AddIfUnique(points, dedupe, new Point2d(maxX, minY));
          AddIfUnique(points, dedupe, new Point2d(maxX, maxY));
          AddIfUnique(points, dedupe, new Point2d(minX, maxY));

          if (first)
          {
            elevation = ext.MinPoint.Z;
            first = false;
          }

          included++;
        }

        if (points.Count < 3)
        {
          ed.WriteMessage("\nNot enough block references with valid extents were selected.");
          return;
        }

        var hull = ComputeConvexHull(points);
        if (hull == null || hull.Count < 3)
        {
          ed.WriteMessage("\nFailed to compute convex hull boundary.");
          return;
        }

        var boundary = BuildPolylineFromPoints(hull, elevation);
        if (boundary == null)
        {
          ed.WriteMessage("\nFailed to create boundary polyline.");
          return;
        }

        if (!model.IsWriteEnabled)
        {
          model.UpgradeOpen();
        }

        boundary.SetDatabaseDefaults();
        boundary.Normal = Vector3d.ZAxis;
        boundary.Closed = true;

        model.AppendEntity(boundary);
        tr.AddNewlyCreatedDBObject(boundary, true);

        tr.Commit();

        ed.WriteMessage($"\nCreated boundary polyline around {included} block reference(s).");
        if (skipped > 0)
        {
          ed.WriteMessage($"\nIgnored {skipped} selection(s) that were not block references in model space or lacked extents.");
        }
        ed.Regen();
      }
    }

    [CommandMethod("BLOCKBOUNDARYGRID", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public void CreateGridUnionBoundaryAroundBlockReferences()
    {
      var doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      var db = doc.Database;
      var ed = doc.Editor;

      SelectionSet selection = null;

      var implied = ed.SelectImplied();
      if (implied.Status == PromptStatus.OK && implied.Value != null && implied.Value.Count > 0)
      {
        selection = implied.Value;
        ed.SetImpliedSelection(Array.Empty<ObjectId>());
      }

      if (selection == null)
      {
        var selOpts = new PromptSelectionOptions
        {
          MessageForAdding = "\nSelect block references in model space: ",
          MessageForRemoval = "\nRemove block references: ",
          AllowDuplicates = false,
          RejectObjectsOnLockedLayers = true
        };

        var filter = new SelectionFilter(new[]
        {
          new TypedValue((int)DxfCode.Start, "INSERT")
        });

        var selRes = ed.GetSelection(selOpts, filter);
        if (selRes.Status != PromptStatus.OK)
        {
          return;
        }

        selection = selRes.Value;
      }

      if (selection == null || selection.Count == 0)
      {
        ed.WriteMessage("\nNo block references selected.");
        return;
      }

      var paddingOptions = new PromptDoubleOptions("\nEnter padding distance: ")
      {
        AllowNegative = false,
        AllowZero = true,
        DefaultValue = 0.0,
        UseDefaultValue = true
      };

      var paddingResult = ed.GetDouble(paddingOptions);
      if (paddingResult.Status != PromptStatus.OK)
      {
        return;
      }

      double padding = paddingResult.Value;

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var modelId = bt[BlockTableRecord.ModelSpace];
        var model = (BlockTableRecord)tr.GetObject(modelId, OpenMode.ForRead);

        var rectangles = new List<(double minX, double minY, double maxX, double maxY)>();

        int included = 0;
        int skipped = 0;
        double elevation = 0.0;
        bool first = true;

        foreach (var id in selection.GetObjectIds())
        {
          if (id == ObjectId.Null) continue;

          var obj = tr.GetObject(id, OpenMode.ForRead, false);
          if (obj is not BlockReference blockRef)
          {
            skipped++;
            continue;
          }

          if (blockRef.OwnerId != modelId)
          {
            skipped++;
            continue;
          }

          var extents = TryGetEntityExtents(blockRef);
          if (extents == null)
          {
            skipped++;
            continue;
          }

          var ext = extents.Value;
          double minX = ext.MinPoint.X - padding;
          double minY = ext.MinPoint.Y - padding;
          double maxX = ext.MaxPoint.X + padding;
          double maxY = ext.MaxPoint.Y + padding;

          rectangles.Add((minX, minY, maxX, maxY));

          if (first)
          {
            elevation = ext.MinPoint.Z;
            first = false;
          }

          included++;
        }

        if (rectangles.Count == 0)
        {
          ed.WriteMessage("\nNo qualifying block references were found in model space.");
          return;
        }

        var loops = BuildGridUnionLoops(rectangles);
        if (loops == null || loops.Count == 0)
        {
          ed.WriteMessage("\nFailed to construct union boundary.");
          return;
        }

        if (loops.Count > 1 && ConnectLoopRectangles(rectangles, loops, padding))
        {
          loops = BuildGridUnionLoops(rectangles);
          if (loops == null || loops.Count == 0)
          {
            ed.WriteMessage("\nFailed to construct union boundary after merging loops.");
            return;
          }
        }

        var outerLoop = loops
          .Select(loop => (Loop: loop, Area: Math.Abs(ComputeSignedArea(loop))))
          .Where(entry => entry.Loop != null && entry.Loop.Count >= 3 && entry.Area > VertexTolerance)
          .OrderByDescending(entry => entry.Area)
          .Select(entry => entry.Loop)
          .FirstOrDefault();

        if (outerLoop == null || outerLoop.Count < 3)
        {
          ed.WriteMessage("\nUnable to determine a valid outer boundary.");
          return;
        }

        EnsureCounterClockwise(outerLoop);
        var boundary = BuildPolylineFromPoints(outerLoop, elevation);
        if (boundary == null)
        {
          ed.WriteMessage("\nFailed to create boundary polyline.");
          return;
        }

        if (!model.IsWriteEnabled)
        {
          model.UpgradeOpen();
        }

        boundary.SetDatabaseDefaults();
        boundary.Normal = Vector3d.ZAxis;
        boundary.Closed = true;

        model.AppendEntity(boundary);
        tr.AddNewlyCreatedDBObject(boundary, true);

        tr.Commit();

        ed.WriteMessage($"\nCreated boundary polyline around {included} block reference(s).");
        if (skipped > 0)
        {
          ed.WriteMessage($"\nIgnored {skipped} selection(s) that were not block references in model space or lacked extents.");
        }
        ed.Regen();
      }
    }

    private static void AddIfUnique(List<Point2d> points, HashSet<Point2d> dedupe, Point2d candidate)
    {
      if (dedupe.Add(candidate))
      {
        points.Add(candidate);
      }
    }

    private static List<Point2d> ComputeConvexHull(List<Point2d> points)
    {
      if (points == null || points.Count < 3)
      {
        return points ?? new List<Point2d>();
      }

      var pts = points.OrderBy(p => p.X).ThenBy(p => p.Y).ToList();

      var lower = new List<Point2d>();
      foreach (var pt in pts)
      {
        while (lower.Count >= 2 && Cross(lower[lower.Count - 2], lower[lower.Count - 1], pt) <= 0)
        {
          lower.RemoveAt(lower.Count - 1);
        }
        lower.Add(pt);
      }

      var upper = new List<Point2d>();
      for (int i = pts.Count - 1; i >= 0; i--)
      {
        var pt = pts[i];
        while (upper.Count >= 2 && Cross(upper[upper.Count - 2], upper[upper.Count - 1], pt) <= 0)
        {
          upper.RemoveAt(upper.Count - 1);
        }
        upper.Add(pt);
      }

      lower.RemoveAt(lower.Count - 1);
      upper.RemoveAt(upper.Count - 1);
      lower.AddRange(upper);

      return lower;
    }

    private static double Cross(Point2d origin, Point2d a, Point2d b)
    {
      return (a.X - origin.X) * (b.Y - origin.Y) - (a.Y - origin.Y) * (b.X - origin.X);
    }

    private static Polyline BuildPolylineFromPoints(List<Point2d> points, double elevation)
    {
      if (points == null || points.Count < 3)
      {
        return null;
      }

      int count = points.Count;
      if (PointsNearlyEqual(points[0], points[count - 1]))
      {
        count--;
      }

      if (count < 3)
      {
        return null;
      }

      var poly = new Polyline(count);
      for (int i = 0; i < count; i++)
      {
        poly.AddVertexAt(i, points[i], 0, 0, 0);
      }

      poly.Closed = true;
      poly.Elevation = elevation;
      poly.Normal = Vector3d.ZAxis;
      return poly;
    }

    private static List<List<Point2d>> BuildGridUnionLoops(List<(double minX, double minY, double maxX, double maxY)> rectangles)
    {
      var result = new List<List<Point2d>>();
      if (rectangles == null || rectangles.Count == 0)
      {
        return result;
      }

      var xValues = new List<double>();
      var yValues = new List<double>();

      foreach (var rect in rectangles)
      {
        xValues.Add(rect.minX);
        xValues.Add(rect.maxX);
        yValues.Add(rect.minY);
        yValues.Add(rect.maxY);
      }

      xValues.Sort();
      yValues.Sort();

      var uniqueX = DeduplicateSorted(xValues);
      var uniqueY = DeduplicateSorted(yValues);

      if (uniqueX.Count < 2 || uniqueY.Count < 2)
      {
        return result;
      }

      var xLookup = BuildIndexLookup(uniqueX);
      var yLookup = BuildIndexLookup(uniqueY);

      var edges = new HashSet<SegmentKey>();

      foreach (var rect in rectangles)
      {
        int xStart = FindIndex(uniqueX, xLookup, rect.minX);
        int xEnd = FindIndex(uniqueX, xLookup, rect.maxX);
        int yStart = FindIndex(uniqueY, yLookup, rect.minY);
        int yEnd = FindIndex(uniqueY, yLookup, rect.maxY);

        if (xStart < 0 || xEnd < 0 || yStart < 0 || yEnd < 0)
        {
          continue;
        }

        for (int xi = xStart; xi < xEnd; xi++)
        {
          double x0 = uniqueX[xi];
          double x1 = uniqueX[xi + 1];

          for (int yi = yStart; yi < yEnd; yi++)
          {
            double y0 = uniqueY[yi];
            double y1 = uniqueY[yi + 1];

            ToggleEdge(edges, new Point2d(x0, y0), new Point2d(x1, y0));
            ToggleEdge(edges, new Point2d(x1, y0), new Point2d(x1, y1));
            ToggleEdge(edges, new Point2d(x1, y1), new Point2d(x0, y1));
            ToggleEdge(edges, new Point2d(x0, y1), new Point2d(x0, y0));
          }
        }
      }

      if (edges.Count == 0)
      {
        return result;
      }

      result.AddRange(ExtractLoopsFromEdges(edges));
      return result;
    }

    private static bool ConnectLoopRectangles(List<(double minX, double minY, double maxX, double maxY)> rectangles, List<List<Point2d>> loops, double padding)
    {
      var bounds = new List<(double minX, double minY, double maxX, double maxY)>();
      var areas = new List<double>();

      foreach (var loop in loops)
      {
        if (loop == null || loop.Count < 3)
        {
          continue;
        }

        double area = Math.Abs(ComputeSignedArea(loop));
        if (area <= VertexTolerance)
        {
          continue;
        }

        bounds.Add(GetLoopBounds(loop));
        areas.Add(area);
      }

      if (bounds.Count <= 1)
      {
        return false;
      }

      int start = 0;
      for (int i = 1; i < areas.Count; i++)
      {
        if (areas[i] > areas[start])
        {
          start = i;
        }
      }

      var connected = new HashSet<int> { start };
      var remaining = new HashSet<int>(Enumerable.Range(0, bounds.Count).Where(i => i != start));
      bool added = false;

      while (remaining.Count > 0)
      {
        double bestDistance = double.PositiveInfinity;
        int bestFrom = -1;
        int bestTo = -1;

        foreach (var from in connected)
        {
          foreach (var to in remaining)
          {
            double distance = DistanceBetweenBoxes(bounds[from], bounds[to]);
            if (distance < bestDistance - VertexTolerance ||
                (Math.Abs(distance - bestDistance) <= VertexTolerance && areas[to] > (bestTo >= 0 ? areas[bestTo] : double.NegativeInfinity)))
            {
              bestDistance = distance;
              bestFrom = from;
              bestTo = to;
            }
          }
        }

        if (bestTo == -1)
        {
          break;
        }

        var connectors = CreateConnectorRectangles(bounds[bestFrom], bounds[bestTo], padding).ToList();
        if (connectors.Count == 0)
        {
          connectors.Add(NormalizeRect(
            Math.Min(bounds[bestFrom].minX, bounds[bestTo].minX),
            Math.Min(bounds[bestFrom].minY, bounds[bestTo].minY),
            Math.Max(bounds[bestFrom].maxX, bounds[bestTo].maxX),
            Math.Max(bounds[bestFrom].maxY, bounds[bestTo].maxY)));
        }

        rectangles.AddRange(connectors);
        added = true;
        connected.Add(bestTo);
        remaining.Remove(bestTo);
      }

      return added;
    }

    private static (double minX, double minY, double maxX, double maxY) GetLoopBounds(List<Point2d> loop)
    {
      double minX = double.PositiveInfinity;
      double minY = double.PositiveInfinity;
      double maxX = double.NegativeInfinity;
      double maxY = double.NegativeInfinity;

      foreach (var pt in loop)
      {
        if (pt.X < minX) minX = pt.X;
        if (pt.Y < minY) minY = pt.Y;
        if (pt.X > maxX) maxX = pt.X;
        if (pt.Y > maxY) maxY = pt.Y;
      }

      return (minX, minY, maxX, maxY);
    }

    private static double DistanceBetweenBoxes((double minX, double minY, double maxX, double maxY) a,
                                               (double minX, double minY, double maxX, double maxY) b)
    {
      double dx = 0.0;
      if (a.maxX < b.minX)
      {
        dx = b.minX - a.maxX;
      }
      else if (b.maxX < a.minX)
      {
        dx = a.minX - b.maxX;
      }

      double dy = 0.0;
      if (a.maxY < b.minY)
      {
        dy = b.minY - a.maxY;
      }
      else if (b.maxY < a.minY)
      {
        dy = a.minY - b.maxY;
      }

      return Math.Max(0.0, dx) + Math.Max(0.0, dy);
    }

    private static IEnumerable<(double minX, double minY, double maxX, double maxY)> CreateConnectorRectangles(
      (double minX, double minY, double maxX, double maxY) a,
      (double minX, double minY, double maxX, double maxY) b,
      double padding)
    {
      var result = new List<(double, double, double, double)>();

      double width = Math.Max(padding, 1.0);
      double halfWidth = width * 0.5;

      double overlapX = Math.Min(a.maxX, b.maxX) - Math.Max(a.minX, b.minX);
      double overlapY = Math.Min(a.maxY, b.maxY) - Math.Max(a.minY, b.minY);

      if (overlapX >= -VertexTolerance)
      {
        double midX = (Math.Max(a.minX, b.minX) + Math.Min(a.maxX, b.maxX)) * 0.5;
        double minX = midX - halfWidth;
        double maxX = midX + halfWidth;
        double minY = Math.Min(a.maxY, b.maxY);
        double maxY = Math.Max(a.minY, b.minY);
        result.Add(NormalizeRect(minX, minY, maxX, maxY));
        return result;
      }

      if (overlapY >= -VertexTolerance)
      {
        double midY = (Math.Max(a.minY, b.minY) + Math.Min(a.maxY, b.maxY)) * 0.5;
        double minY = midY - halfWidth;
        double maxY = midY + halfWidth;
        double minX = Math.Min(a.maxX, b.maxX);
        double maxX = Math.Max(a.minX, b.minX);
        result.Add(NormalizeRect(minX, minY, maxX, maxY));
        return result;
      }

      double centerAX = (a.minX + a.maxX) * 0.5;
      double centerAY = (a.minY + a.maxY) * 0.5;
      double centerBX = (b.minX + b.maxX) * 0.5;
      double centerBY = (b.minY + b.maxY) * 0.5;

      double minXVertical = Math.Min(centerAX, centerBX) - halfWidth;
      double maxXVertical = Math.Min(centerAX, centerBX) + halfWidth;
      double minYVertical = Math.Min(centerAY, centerBY);
      double maxYVertical = Math.Max(centerAY, centerBY);
      result.Add(NormalizeRect(minXVertical, minYVertical, maxXVertical, maxYVertical));

      double minXHorizontal = Math.Min(centerAX, centerBX);
      double maxXHorizontal = Math.Max(centerAX, centerBX);
      double minYHorizontal = centerBY - halfWidth;
      double maxYHorizontal = centerBY + halfWidth;
      result.Add(NormalizeRect(minXHorizontal, minYHorizontal, maxXHorizontal, maxYHorizontal));

      return result;
    }

    private static (double minX, double minY, double maxX, double maxY) NormalizeRect(double minX, double minY, double maxX, double maxY)
    {
      if (maxX < minX)
      {
        (minX, maxX) = (maxX, minX);
      }

      if (maxY < minY)
      {
        (minY, maxY) = (maxY, minY);
      }

      double minSize = VertexTolerance * 10.0;

      if (maxX - minX < minSize)
      {
        double mid = (minX + maxX) * 0.5;
        minX = mid - minSize * 0.5;
        maxX = mid + minSize * 0.5;
      }

      if (maxY - minY < minSize)
      {
        double mid = (minY + maxY) * 0.5;
        minY = mid - minSize * 0.5;
        maxY = mid + minSize * 0.5;
      }

      return (minX, minY, maxX, maxY);
    }

    private static void EnsureCounterClockwise(List<Point2d> points)
    {
      if (points == null || points.Count < 3)
      {
        return;
      }

      if (ComputeSignedArea(points) < 0.0)
      {
        points.Reverse();
      }
    }

    private static double ComputeSignedArea(List<Point2d> points)
    {
      double area = 0.0;
      if (points == null || points.Count < 3)
      {
        return 0.0;
      }

      int count = points.Count;
      for (int i = 0; i < count; i++)
      {
        var a = points[i];
        var b = points[(i + 1) % count];
        area += (a.X * b.Y) - (b.X * a.Y);
      }

      return area * 0.5;
    }

    private static List<double> DeduplicateSorted(List<double> values)
    {
      var result = new List<double>();
      if (values == null || values.Count == 0)
      {
        return result;
      }

      double last = values[0];
      result.Add(last);

      for (int i = 1; i < values.Count; i++)
      {
        double current = values[i];
        if (Math.Abs(current - last) > VertexTolerance)
        {
          result.Add(current);
          last = current;
        }
      }

      return result;
    }

    private static Dictionary<long, int> BuildIndexLookup(List<double> coords)
    {
      var lookup = new Dictionary<long, int>();
      for (int i = 0; i < coords.Count; i++)
      {
        lookup[QuantizeCoordinate(coords[i])] = i;
      }
      return lookup;
    }

    private static int FindIndex(List<double> coords, Dictionary<long, int> lookup, double value)
    {
      if (lookup.TryGetValue(QuantizeCoordinate(value), out int idx))
      {
        return idx;
      }

      for (int i = 0; i < coords.Count; i++)
      {
        if (Math.Abs(coords[i] - value) <= VertexTolerance)
        {
          return i;
        }
      }

      return -1;
    }

    private static long QuantizeCoordinate(double value)
    {
      return (long)Math.Round(value / VertexTolerance, MidpointRounding.AwayFromZero);
    }

    private static void ToggleEdge(HashSet<SegmentKey> edges, Point2d a, Point2d b)
    {
      var key = new SegmentKey(a, b);
      if (!edges.Add(key))
      {
        edges.Remove(key);
      }
    }

    private static List<List<Point2d>> ExtractLoopsFromEdges(IEnumerable<SegmentKey> edges)
    {
      var comparer = new Point2dEqualityComparer(VertexTolerance);
      var adjacency = new Dictionary<Point2d, List<Point2d>>(comparer);

      foreach (var edge in edges)
      {
        AddNeighbor(adjacency, edge.Start, edge.End, comparer);
        AddNeighbor(adjacency, edge.End, edge.Start, comparer);
      }

      var loops = new List<List<Point2d>>();

      while (adjacency.Count > 0)
      {
        var start = GetNextStartPoint(adjacency);
        var loop = TraceLoop(adjacency, start, comparer);
        if (loop != null && loop.Count >= 3)
        {
          loops.Add(loop);
        }
      }

      return loops;
    }

    private static Point2d GetNextStartPoint(Dictionary<Point2d, List<Point2d>> adjacency)
    {
      Point2d best = default;
      bool initialized = false;

      foreach (var kvp in adjacency)
      {
        var pt = kvp.Key;
        if (!initialized || pt.Y < best.Y - VertexTolerance ||
            (Math.Abs(pt.Y - best.Y) <= VertexTolerance && pt.X < best.X))
        {
          best = pt;
          initialized = true;
        }
      }

      return best;
    }

    private static List<Point2d> TraceLoop(Dictionary<Point2d, List<Point2d>> adjacency, Point2d start, Point2dEqualityComparer comparer)
    {
      var loop = new List<Point2d>();
      var previous = new Point2d(start.X - 1.0, start.Y);
      var current = start;

      int guard = 0;
      while (guard++ < 100000)
      {
        loop.Add(current);

        if (!adjacency.TryGetValue(current, out var neighbors) || neighbors.Count == 0)
        {
          break;
        }

        var next = SelectNextNeighbor(previous, current, neighbors);

        RemoveNeighbor(adjacency, current, next, comparer);
        RemoveNeighbor(adjacency, next, current, comparer);

        previous = current;
        current = next;

        if (comparer.Equals(current, start))
        {
          loop.Add(start);
          break;
        }
      }

      return loop;
    }

    private static Point2d SelectNextNeighbor(Point2d previous, Point2d current, List<Point2d> candidates)
    {
      if (candidates == null || candidates.Count == 0)
      {
        return current;
      }

      if (candidates.Count == 1)
      {
        return candidates[0];
      }

      var prevDir = current - previous;
      if (prevDir.Length <= VertexTolerance)
      {
        prevDir = Vector2d.XAxis;
      }

      Point2d best = candidates[0];
      double bestAngle = double.NegativeInfinity;
      double bestDot = double.NegativeInfinity;

      foreach (var candidate in candidates)
      {
        var dir = candidate - current;
        if (dir.Length <= VertexTolerance)
        {
          continue;
        }

        double cross = (prevDir.X * dir.Y) - (prevDir.Y * dir.X);
        double dot = (prevDir.X * dir.X) + (prevDir.Y * dir.Y);
        double angle = Math.Atan2(cross, dot);

        if (angle > bestAngle + 1e-12 || (Math.Abs(angle - bestAngle) <= 1e-12 && dot > bestDot))
        {
          best = candidate;
          bestAngle = angle;
          bestDot = dot;
        }
      }

      return best;
    }

    private static void AddNeighbor(Dictionary<Point2d, List<Point2d>> adjacency, Point2d point, Point2d neighbor, Point2dEqualityComparer comparer)
    {
      if (!adjacency.TryGetValue(point, out var neighbors))
      {
        neighbors = new List<Point2d>();
        adjacency[point] = neighbors;
      }

      for (int i = 0; i < neighbors.Count; i++)
      {
        if (comparer.Equals(neighbors[i], neighbor))
        {
          return;
        }
      }

      neighbors.Add(neighbor);
    }

    private static void RemoveNeighbor(Dictionary<Point2d, List<Point2d>> adjacency, Point2d point, Point2d neighbor, Point2dEqualityComparer comparer)
    {
      if (!adjacency.TryGetValue(point, out var neighbors))
      {
        return;
      }

      for (int i = 0; i < neighbors.Count; i++)
      {
        if (comparer.Equals(neighbors[i], neighbor))
        {
          neighbors.RemoveAt(i);
          break;
        }
      }

      if (neighbors.Count == 0)
      {
        adjacency.Remove(point);
      }
    }
    private static bool PointsNearlyEqual(Point2d a, Point2d b)
    {
      return Math.Abs(a.X - b.X) <= VertexTolerance && Math.Abs(a.Y - b.Y) <= VertexTolerance;
    }

    private static Extents3d? TryGetEntityExtents(Entity entity)
    {
      if (entity == null)
      {
        return null;
      }

      try
      {
        return entity.GeometricExtents;
      }
      catch
      {
        try
        {
          return entity.Bounds;
        }
        catch
        {
          return null;
        }
      }
    }

    private sealed class Point2dEqualityComparer : IEqualityComparer<Point2d>
    {
      private readonly double _tolerance;

      public Point2dEqualityComparer(double tolerance)
      {
        _tolerance = Math.Max(tolerance, 1e-12);
      }

      public bool Equals(Point2d x, Point2d y)
      {
        return Math.Abs(x.X - y.X) <= _tolerance && Math.Abs(x.Y - y.Y) <= _tolerance;
      }

      public int GetHashCode(Point2d obj)
      {
        long qx = (long)Math.Round(obj.X / _tolerance, MidpointRounding.AwayFromZero);
        long qy = (long)Math.Round(obj.Y / _tolerance, MidpointRounding.AwayFromZero);
        unchecked
        {
          int hash = 17;
          hash = (hash * 31) + qx.GetHashCode();
          hash = (hash * 31) + qy.GetHashCode();
          return hash;
        }
      }
    }

    private readonly struct SegmentKey
    {
      public SegmentKey(Point2d a, Point2d b)
      {
        Start = a;
        End = b;

        var keyA = (X: QuantizeCoordinate(a.X), Y: QuantizeCoordinate(a.Y));
        var keyB = (X: QuantizeCoordinate(b.X), Y: QuantizeCoordinate(b.Y));

        if (keyA.X < keyB.X || (keyA.X == keyB.X && keyA.Y <= keyB.Y))
        {
          KeyA = keyA;
          KeyB = keyB;
        }
        else
        {
          KeyA = keyB;
          KeyB = keyA;
          Start = b;
          End = a;
        }
      }

      private (long X, long Y) KeyA { get; }
      private (long X, long Y) KeyB { get; }

      public Point2d Start { get; }
      public Point2d End { get; }

      public override bool Equals(object obj)
      {
        return obj is SegmentKey other &&
               KeyA.X == other.KeyA.X && KeyA.Y == other.KeyA.Y &&
               KeyB.X == other.KeyB.X && KeyB.Y == other.KeyB.Y;
      }

      public override int GetHashCode()
      {
        unchecked
        {
          int hash = 17;
          hash = (hash * 31) + KeyA.X.GetHashCode();
          hash = (hash * 31) + KeyA.Y.GetHashCode();
          hash = (hash * 31) + KeyB.X.GetHashCode();
          hash = (hash * 31) + KeyB.Y.GetHashCode();
          return hash;
        }
      }
    }
  }
}
