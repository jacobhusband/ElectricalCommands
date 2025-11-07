using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        [CommandMethod("VP2PL", CommandFlags.Modal)]
        public static void ViewportToPolyline_AllLayouts()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                using (doc.LockDocument())
                {
                    var modelPolys = new List<List<Point3d>>();

                    // ---- Enumerate ALL paper-space layouts via Layout Dictionary ----
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                        foreach (DBDictionaryEntry kv in layoutDict)
                        {
                            var layout = (Layout)tr.GetObject(kv.Value, OpenMode.ForRead);
                            if (layout == null) continue;
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) continue;

                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                            var tbUnion = GetTitleBlockUnionExtentsInPaper_Robust(btr, tr);

                            var vpIds = new List<ObjectId>();
                            foreach (ObjectId id in btr)
                            {
                                if (!id.IsValid) continue;
                                if (id.ObjectClass.DxfName != "VIEWPORT") continue;
                                var vp = tr.GetObject(id, OpenMode.ForRead) as Viewport;
                                if (vp == null || vp.Number == 1) continue;
                                vpIds.Add(id);
                            }
                            if (vpIds.Count == 0) continue;

                            int added = 0;
                            bool haveTb = tbUnion.HasValue;

                            foreach (var vpId in vpIds)
                            {
                                var vp = (Viewport)tr.GetObject(vpId, OpenMode.ForRead);

                                bool include = true;
                                if (haveTb)
                                {
                                    var vpEx = GetViewportRectExtentsInPaper(vp);
                                    include = ExtentsOverlap2d(
                                        ExpandExtents2(tbUnion.Value, margin: 2.0),
                                        ExpandExtents2(vpEx, margin: 0.0)
                                    );
                                }
                                if (!include) continue;

                                var ps2d = GetViewportBoundaryPointsInPaper(vp, tr);
                                if (ps2d == null || ps2d.Count < 3) continue;

                                var ms3d = PaperPolygonToModel_Correct(vp, ps2d);
                                if (ms3d != null && ms3d.Count >= 3)
                                {
                                    modelPolys.Add(ms3d);
                                    added++;
                                }
                            }

                            if (added == 0)
                            {
                                foreach (var vpId in vpIds)
                                {
                                    var vp = (Viewport)tr.GetObject(vpId, OpenMode.ForRead);
                                    var ps2d = GetViewportBoundaryPointsInPaper(vp, tr);
                                    if (ps2d == null || ps2d.Count < 3) continue;

                                    var ms3d = PaperPolygonToModel_Correct(vp, ps2d);
                                    if (ms3d != null && ms3d.Count >= 3)
                                        modelPolys.Add(ms3d);
                                }
                            }
                        }

                        tr.Commit();
                    }

                    SwitchToModelSpaceViewSafe(db, ed);
                    var msId = SymbolUtilityServices.GetBlockModelSpaceId(db);

                    if (modelPolys.Count == 0)
                    {
                        ed.WriteMessage("\nNo eligible viewport regions found on any paper space layout. Erasing all objects in Model Space.");
                        // Erase everything by providing an empty "keep" set
                        int erased = EraseEntitiesExcept(db, msId, new HashSet<ObjectId>());
                        if (erased > 0)
                        {
                            ed.WriteMessage($"\nErased all {erased} object(s) from Model Space.");
                        }
                        else
                        {
                            ed.WriteMessage("\nModel Space was already empty.");
                        }
                        return;
                    }

                    // ---- Select + Keep in Model space, erase everything else ----
                    var keepIds = new HashSet<ObjectId>();
                    var originalUcs = ed.CurrentUserCoordinateSystem;
                    try
                    {
                        ed.CurrentUserCoordinateSystem = Matrix3d.Identity;

                        foreach (var poly in modelPolys)
                        {
                            var coll = new Point3dCollection(poly.ToArray());
                            var res = ed.SelectCrossingPolygon(coll);
                            if (res.Status == PromptStatus.OK)
                            {
                                foreach (var oid in res.Value.GetObjectIds())
                                    keepIds.Add(oid);
                            }
                        }
                    }
                    finally
                    {
                        ed.CurrentUserCoordinateSystem = originalUcs;
                    }

                    if (keepIds.Count == 0)
                    {
                        ed.WriteMessage("\nNothing found inside any viewport regions. Erasing all objects in Model Space.");
                        int erased = EraseEntitiesExcept(db, msId, new HashSet<ObjectId>());
                        ed.WriteMessage($"\nErased all {erased} object(s) from Model Space.");
                        return;
                    }

                    int erasedCount = EraseEntitiesExcept(db, msId, keepIds);
                    ed.WriteMessage($"\nVP2PL: Kept {keepIds.Count} object(s) inside/crossing all viewport regions; erased {erasedCount} others.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nVP2PL failed: {ex.Message}");
            }
        }

        // -------------------- Updated / New Helpers --------------------

        // Erase all entities in a space except those in keep-set. Returns count erased.
        private static int EraseEntitiesExcept(Database db, ObjectId spaceId, HashSet<ObjectId> keep)
        {
            int erased = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                var toErase = new List<Entity>();
                var layersToUnlock = new Dictionary<ObjectId, bool>();

                // First pass: identify what to erase and what layers to unlock
                foreach (ObjectId id in btr)
                {
                    if (!id.IsValid || keep.Contains(id)) continue;

                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null) continue;

                    toErase.Add(ent);

                    if (!layersToUnlock.ContainsKey(ent.LayerId))
                    {
                        var layer = (LayerTableRecord)tr.GetObject(ent.LayerId, OpenMode.ForRead);
                        if (layer.IsLocked)
                        {
                            layersToUnlock[ent.LayerId] = true;
                        }
                    }
                }

                // Unlock layers
                foreach (var layerId in layersToUnlock.Keys)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    layer.IsLocked = false;
                }

                // Erase entities
                foreach (var ent in toErase)
                {
                    try
                    {
                        if (!ent.IsErased)
                        {
                            ent.UpgradeOpen();
                            ent.Erase();
                            erased++;
                        }
                    }
                    catch { /* keep going */ }
                }

                // Re-lock layers
                foreach (var layerId in layersToUnlock.Keys)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    layer.IsLocked = true;
                }

                tr.Commit();
            }
            return erased;
        }

        private static void SwitchToModelSpaceViewSafe(Database db, Editor ed)
        {
            try
            {
                var modelId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                if (db.CurrentSpaceId != modelId)
                {
                    try { ed.SwitchToModelSpace(); }
                    catch { try { Application.SetSystemVariable("TILEMODE", 1); } catch { } }
                }
            }
            catch { }
        }

        // More robust TB extents finder: named block refs OR largest closed polyline. Returns null if nothing plausible.
        private static Extents3d? GetTitleBlockUnionExtentsInPaper_Robust(BlockTableRecord layoutBtr, Transaction tr)
        {
            var named = new List<Extents3d>();
            var allBlocks = new List<(Extents3d ex, double area, string name)>();

            foreach (ObjectId id in layoutBtr)
            {
                if (!id.IsValid) continue;

                // Block refs
                if (id.ObjectClass.DxfName == "INSERT")
                {
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br != null)
                    {
                        try
                        {
                            var def = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            var name = def?.Name ?? string.Empty;
                            var ex = br.GeometricExtents;
                            double area = Math.Max(0, (ex.MaxPoint.X - ex.MinPoint.X)) * Math.Max(0, (ex.MaxPoint.Y - ex.MinPoint.Y));
                            if (area <= 1e-9) continue;

                            allBlocks.Add((ex, area, name));

                            var n = (name ?? string.Empty).ToUpperInvariant();
                            if (n.Contains("X-TB") || n.Contains(" TB") || n.StartsWith("TB") || n.Contains("TITLE") || n.Contains("BORDER"))
                                named.Add(ex);
                        }
                        catch { }
                    }
                }
            }

            // Prefer named matches (union all)
            if (named.Count > 0)
                return UnionExtents(named);

            // Fallback #1: largest block ref by extents area
            if (allBlocks.Count > 0)
            {
                var best = allBlocks.OrderByDescending(b => b.area).First().ex;
                return best;
            }

            // Fallback #2: largest closed polyline by extents area
            Extents3d? bestPlEx = null;
            double bestA = 0.0;
            foreach (ObjectId id in layoutBtr)
            {
                if (!id.IsValid) continue;
                var pl = tr.GetObject(id, OpenMode.ForRead) as Polyline;
                if (pl == null || !pl.Closed) continue;
                try
                {
                    var ex = pl.GeometricExtents;
                    double area = Math.Max(0, (ex.MaxPoint.X - ex.MinPoint.X)) * Math.Max(0, (ex.MaxPoint.Y - ex.MinPoint.Y));
                    if (area > bestA) { bestA = area; bestPlEx = ex; }
                }
                catch { }
            }
            return bestPlEx;
        }

        private static Extents3d UnionExtents(List<Extents3d> boxes)
        {
            var ex0 = boxes[0];
            double minX = ex0.MinPoint.X, minY = ex0.MinPoint.Y, minZ = ex0.MinPoint.Z;
            double maxX = ex0.MaxPoint.X, maxY = ex0.MaxPoint.Y, maxZ = ex0.MaxPoint.Z;

            for (int i = 1; i < boxes.Count; i++)
            {
                var ex = boxes[i];
                minX = Math.Min(minX, ex.MinPoint.X);
                minY = Math.Min(minY, ex.MinPoint.Y);
                minZ = Math.Min(minZ, ex.MinPoint.Z);
                maxX = Math.Max(maxX, ex.MaxPoint.X);
                maxY = Math.Max(maxY, ex.MaxPoint.Y);
                maxZ = Math.Max(maxZ, ex.MaxPoint.Z);
            }
            return new Extents3d(new Point3d(minX, minY, minZ), new Point3d(maxX, maxY, maxZ));
        }

        // Axis-aligned rect overlap in Paper (Z ignored)
        private static bool ExtentsOverlap2d(Extents3d a, Extents3d b)
        {
            return !(a.MaxPoint.X < b.MinPoint.X || a.MinPoint.X > b.MaxPoint.X ||
                     a.MaxPoint.Y < b.MinPoint.Y || a.MinPoint.Y > b.MaxPoint.Y);
        }

        private static Extents3d ExpandExtents2(Extents3d ex, double margin)
        {
            return new Extents3d(
                new Point3d(ex.MinPoint.X - margin, ex.MinPoint.Y - margin, ex.MinPoint.Z),
                new Point3d(ex.MaxPoint.X + margin, ex.MaxPoint.Y + margin, ex.MaxPoint.Z)
            );
        }

        private static Extents3d GetViewportRectExtentsInPaper(Viewport vp)
        {
            double w = vp.Width, h = vp.Height;
            var c = vp.CenterPoint;
            return new Extents3d(
                new Point3d(c.X - w / 2.0, c.Y - h / 2.0, 0.0),
                new Point3d(c.X + w / 2.0, c.Y + h / 2.0, 0.0)
            );
        }

        // PAPER-SPACE boundary (rect or clipped)
        private static List<Point2d> GetViewportBoundaryPointsInPaper(Viewport vp, Transaction tr)
        {
            if (vp.NonRectClipOn && !vp.NonRectClipEntityId.IsNull)
            {
                var clipEnt = tr.GetObject(vp.NonRectClipEntityId, OpenMode.ForRead) as Entity;
                if (clipEnt != null)
                {
                    var pl = PolylineFromClipEntity(clipEnt, tr);
                    if (pl != null)
                        return SamplePolyline2d(pl, arcSegsPerQuarter: 12);
                }
            }

            double w = vp.Width, h = vp.Height;
            if (w <= 0 || h <= 0) return null;

            var c = vp.CenterPoint;
            double minX = c.X - w / 2.0, maxX = c.X + w / 2.0;
            double minY = c.Y - h / 2.0, maxY = c.Y + h / 2.0;

            return new List<Point2d>
            {
                new Point2d(minX, minY),
                new Point2d(maxX, minY),
                new Point2d(maxX, maxY),
                new Point2d(minX, maxY)
            };
        }

        // Correct PS→MS mapping (accounts for ViewCenter, ViewHeight/Height, TwistAngle)
        private static List<Point3d> PaperPolygonToModel_Correct(Viewport vp, List<Point2d> ps)
        {
            var ms = new List<Point3d>(ps.Count);

            var z = vp.ViewDirection.GetNormal();
            var x0 = (Math.Abs(z.DotProduct(Vector3d.ZAxis)) > 0.999999)
                        ? Vector3d.XAxis
                        : z.GetPerpendicularVector().GetNormal();
            var y0 = z.CrossProduct(x0).GetNormal();

            double twist = vp.TwistAngle;
            var x = x0 * Math.Cos(twist) + y0 * Math.Sin(twist);
            var y = (-x0) * Math.Sin(twist) + y0 * Math.Cos(twist);

            double muPerPu = vp.ViewHeight / vp.Height;
            var target = vp.ViewTarget;
            var vc = vp.ViewCenter; // DCS coords at viewport center

            foreach (var p in ps)
            {
                double dxPS = p.X - vp.CenterPoint.X;
                double dyPS = p.Y - vp.CenterPoint.Y;

                double u = (vc.X + dxPS * muPerPu);
                double v = (vc.Y + dyPS * muPerPu);

                var pt = target + x.MultiplyBy(u) + y.MultiplyBy(v);
                ms.Add(pt);
            }
            return ms;
        }

        // ---- clip entity → polyline helpers (no debug drawing) ----

        private static Polyline PolylineFromClipEntity(Entity ent, Transaction tr)
        {
            switch (ent)
            {
                case Polyline p2d: return ClonePolyline(p2d);
                case Polyline2d p2: return FromPolyline2d(p2, tr);
                case Circle c: return FromCircle(c);
                case Ellipse e: return FromEllipseApprox(e, 64);
                case Spline s: return FromSplineApprox(s, 128);
                default: return FromExplodeLinesArcs(ent);
            }
        }

        private static Polyline ClonePolyline(Polyline src)
        {
            var dst = new Polyline(src.NumberOfVertices);
            for (int i = 0; i < src.NumberOfVertices; i++)
            {
                var pt = src.GetPoint2dAt(i);
                double bulge = src.GetBulgeAt(i);
                dst.AddVertexAt(i, pt, bulge, 0, 0);
            }
            dst.Closed = src.Closed;
            return dst;
        }

        private static Polyline FromPolyline2d(Polyline2d p2, Transaction tr)
        {
            var verts = new List<(Point2d pt, double bulge)>();
            foreach (ObjectId vId in p2)
            {
                var v = (Vertex2d)tr.GetObject(vId, OpenMode.ForRead);
                verts.Add((new Point2d(v.Position.X, v.Position.Y), v.Bulge));
            }
            if (verts.Count < 2) return null;

            var pl = new Polyline(verts.Count);
            for (int i = 0; i < verts.Count; i++)
                pl.AddVertexAt(i, verts[i].pt, verts[i].bulge, 0, 0);

            pl.Closed = p2.Closed || verts[0].pt.GetDistanceTo(verts[verts.Count - 1].pt) <= 1e-9;
            return pl;
        }

        private static Polyline FromCircle(Circle c)
        {
            var center = c.Center;
            double r = c.Radius;
            if (r <= 0) return null;

            var p0 = new Point2d(center.X + r, center.Y);
            var p1 = new Point2d(center.X, center.Y + r);
            var p2 = new Point2d(center.X - r, center.Y);
            var p3 = new Point2d(center.X, center.Y - r);

            double bulge90 = Math.Tan(Math.PI / 8.0);

            var pl = new Polyline(4);
            pl.AddVertexAt(0, p0, bulge90, 0, 0);
            pl.AddVertexAt(1, p1, bulge90, 0, 0);
            pl.AddVertexAt(2, p2, bulge90, 0, 0);
            pl.AddVertexAt(3, p3, bulge90, 0, 0);
            pl.Closed = true;
            return pl;
        }

        private static Polyline FromEllipseApprox(Ellipse e, int segments)
        {
            if (segments < 8) segments = 8;
            var curve = (Curve)e;
            double t0 = curve.StartParam, t1 = curve.EndParam;

            var pl = new Polyline(segments);
            for (int i = 0; i < segments; i++)
            {
                double t = t0 + (t1 - t0) * (double)i / (double)segments;
                var p = curve.GetPointAtParameter(t);
                pl.AddVertexAt(i, new Point2d(p.X, p.Y), 0, 0, 0);
            }
            var pend = curve.GetPointAtParameter(t1);
            pl.AddVertexAt(segments, new Point2d(pend.X, pend.Y), 0, 0, 0);
            pl.Closed = true;
            return pl;
        }

        private static Polyline FromSplineApprox(Spline s, int segments)
        {
            if (segments < 8) segments = 8;
            var curve = (Curve)s;
            double t0 = curve.StartParam, t1 = curve.EndParam;

            var pl = new Polyline(segments + 1);
            for (int i = 0; i <= segments; i++)
            {
                double t = t0 + (t1 - t0) * (double)i / (double)segments;
                var p = curve.GetPointAtParameter(t);
                pl.AddVertexAt(i, new Point2d(p.X, p.Y), 0, 0, 0);
            }
            pl.Closed = s.Closed;
            return pl;
        }

        private static double BulgeFromArc(Arc arc)
        {
            double r = arc.Radius;
            if (r <= 0) return 0.0;
            double theta = arc.Length / r;

            Vector3d v1 = arc.StartPoint - arc.Center;
            Vector3d v2 = arc.EndPoint - arc.Center;
            double sign = Math.Sign(arc.Normal.DotProduct(v1.CrossProduct(v2)));

            double bulge = Math.Tan(theta / 4.0);
            if (sign < 0) bulge = -bulge;
            return bulge;
        }

        private static Polyline FromExplodeLinesArcs(Entity e)
        {
            try
            {
                using (var res = new DBObjectCollection())
                {
                    e.Explode(res);
                    var segments = new List<(Point3d s, Point3d e, double bulge)>();
                    foreach (DBObject dbo in res)
                    {
                        if (dbo is Line ln) segments.Add((ln.StartPoint, ln.EndPoint, 0.0));
                        else if (dbo is Arc arc) segments.Add((arc.StartPoint, arc.EndPoint, BulgeFromArc(arc)));
                        dbo.Dispose();
                    }
                    if (segments.Count == 0) return null;

                    var chain = StitchSegments(segments);
                    if (chain == null || chain.Count < 2) return null;

                    var pl = new Polyline(chain.Count);
                    for (int i = 0; i < chain.Count; i++)
                    {
                        var node = chain[i];
                        pl.AddVertexAt(i, new Point2d(node.s.X, node.s.Y), node.bulge, 0, 0);
                    }
                    if (chain[0].s.DistanceTo(chain[chain.Count-1].e) < 1e-6) pl.Closed = true;
                    return pl;
                }
            }
            catch { return null; }
        }

        private static List<(Point3d s, Point3d e, double bulge)> StitchSegments(List<(Point3d s, Point3d e, double bulge)> segs)
        {
            if (segs.Count == 0) return null;
            var chain = new List<(Point3d s, Point3d e, double bulge)> { segs[0] };
            segs.RemoveAt(0);

            const double tol = 1e-6;
            while (segs.Count > 0)
            {
                var last = chain[chain.Count - 1];
                int found = -1;
                bool reverse = false;

                for (int i = 0; i < segs.Count; i++)
                {
                    if (last.e.DistanceTo(segs[i].s) < tol) { found = i; reverse = false; break; }
                    if (last.e.DistanceTo(segs[i].e) < tol) { found = i; reverse = true; break; }
                }
                if (found < 0) break;

                var seg = segs[found];
                segs.RemoveAt(found);
                chain.Add(reverse ? (seg.e, seg.s, -seg.bulge) : seg);
            }
            return chain;
        }

        private static List<Point2d> SamplePolyline2d(Polyline pl, int arcSegsPerQuarter = 12)
        {
            var pts = new List<Point2d>();
            int n = pl.NumberOfVertices;
            if (n == 0) return pts;

            for (int i = 0; i < n; i++)
            {
                var a = pl.GetPoint2dAt(i);
                var b = pl.GetPoint2dAt((i + 1) % n);
                double bulge = pl.GetBulgeAt(i);

                pts.Add(a);

                if (Math.Abs(bulge) > 1e-12)
                {
                    foreach (var q in SampleBulge(a, b, bulge, Math.Max(2, (int)Math.Ceiling(Math.Abs(4 * Math.Atan(bulge)) / (Math.PI / (2.0 * arcSegsPerQuarter))))))
                        pts.Add(q);
                }
            }
            return pts;
        }

        private static IEnumerable<Point2d> SampleBulge(Point2d a, Point2d b, double bulge, int segs)
        {
            Vector2d v = b - a;
            double L = v.Length;
            if (L < 1e-12) yield break;

            double theta = 4.0 * Math.Atan(bulge);
            double half = theta / 2.0;
            double d = (L / 2.0) / Math.Tan(half);

            Vector2d t = v / L;
            Vector2d n = new Vector2d(-t.Y, t.X);
            if (bulge < 0) n = -n;

            var mid = new Point2d((a.X + b.X) * 0.5, (a.Y + b.Y) * 0.5);
            var cen = new Point2d(mid.X + n.X * d, mid.Y + n.Y * d);

            double ang0 = Math.Atan2(a.Y - cen.Y, a.X - cen.X);
            double R = (new Vector2d(a.X - cen.X, a.Y - cen.Y)).Length;

            for (int i = 1; i < segs; i++)
            {
                double ang = ang0 + (theta * i) / segs;
                yield return new Point2d(cen.X + Math.Cos(ang) * R, cen.Y + Math.Sin(ang) * R);
            }
        }
    }
}