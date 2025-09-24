using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsSystem;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        // Find the titleblock region, preselect it and everything inside, then erase everything else in Model Space
        [CommandMethod("KEEPONLYTITLEBLOCKMS", CommandFlags.Modal)]
        public static void KeepOnlyTitleBlockInModelSpace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                var modelId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                if (db.CurrentSpaceId != modelId)
                {
                    try { ed.SwitchToModelSpace(); } catch { try { Application.SetSystemVariable("TILEMODE", 1); } catch { } }
                }
            }
            catch { }

            string[] layerTokens = { "TITLE", "TBLK", "BORDER", "SHEET", "FRAME" };
            var attrTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "SHEET", "SHEETNO", "SHEET_NO", "SHEETNUMBER", "DWG", "DWG_NO",
                "DRAWN", "CHECKED", "APPROVED", "PROJECT", "CLIENT", "SCALE", "DATE",
                "REV", "REVISION", "TITLE"
            };

            HashSet<ObjectId> keepIds = null;
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    // Collect ATTDEF points for scoring
                    var attDefs = new List<(Point3d Pos, string Tag)>();
                    foreach (ObjectId id in ms)
                    {
                        if (!id.IsValid) continue;
                        DBObject dbo = tr.GetObject(id, OpenMode.ForRead, false);
                        if (dbo is AttributeDefinition ad && !ad.Invisible)
                            attDefs.Add((ad.Position, ad.Tag));
                    }

                    // Candidate container
                    var candidates = new List<(Extents3d Ext, double Score, double Angle, Point3d[] Poly, ObjectId[] Boundary)>();

                    // 1) Clean 4-vertex closed polylines
                    foreach (ObjectId id in ms)
                    {
                        if (!id.IsValid) continue;
                        if (!(tr.GetObject(id, OpenMode.ForRead, false) is Entity ent)) continue;

                        var ext = TryGetExtents(ent);
                        if (ext == null) continue;
                        double score = 0.0; double angle = 0.0;

                        if (layerTokens.Any(t => ent.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0))
                            score += 0.8;

                        if (ent is Polyline pl && pl.Closed && pl.NumberOfVertices == 4)
                        {
                            var (ok, w, h, ang) = TryRectInfo(pl);
                            if (!ok) continue;
                            angle = ang;
                            double ratio = Math.Max(w, h) / Math.Max(1e-9, Math.Min(w, h));
                            if (IsSheetRatio(ratio)) score += 3.0;
                            score += 0.7;

                            if (attDefs.Count > 0)
                            {
                                int inside = 0;
                                foreach (var a in attDefs) if (PointInExtents2D(a.Pos, ext.Value)) inside++;
                                score += Math.Min(2.0, inside * 0.3);
                                int tagHits = attDefs.Count(a => PointInExtents2D(a.Pos, ext.Value) && attrTags.Contains(a.Tag));
                                score += Math.Min(2.0, tagHits * 0.5);
                            }

                            // Build polygon from polyline vertices
                            var pts = new Point3d[4];
                            for (int i = 0; i < 4; i++)
                            {
                                var p2 = pl.GetPoint2dAt(i);
                                pts[i] = new Point3d(p2.X, p2.Y, 0);
                            }

                            // size bias
                            var s = ext.Value.MaxPoint - ext.Value.MinPoint;
                            double area = Math.Abs(s.X * s.Y);
                            if (area > 1) score += Math.Log10(area);

                            candidates.Add((ext.Value, score, angle, pts, new[] { ent.ObjectId }));
                        }
                    }

                    // 2) Rectangles made of four Lines
                    var lines = new List<(Line L, Point2d P0, Point2d P1, Vector2d V, double Len, string Layer)>();
                    foreach (ObjectId id in ms)
                    {
                        if (!id.IsValid) continue;
                        if (tr.GetObject(id, OpenMode.ForRead, false) is Line ln)
                        {
                            var p0 = new Point2d(ln.StartPoint.X, ln.StartPoint.Y);
                            var p1 = new Point2d(ln.EndPoint.X, ln.EndPoint.Y);
                            var v = p1 - p0;
                            double len = v.Length;
                            if (len > 1e-6) lines.Add((ln, p0, p1, v / len, len, ln.Layer));
                        }
                    }
                    if (lines.Count >= 4)
                    {
                        var hinted = lines.Where(l => layerTokens.Any(t => l.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();
                        var poolSet = new HashSet<ObjectId>(hinted.Select(h => h.L.ObjectId));
                        var pool = new List<(Line L, Point2d P0, Point2d P1, Vector2d V, double Len, string Layer)>(hinted);
                        foreach (var x in lines.OrderByDescending(x => x.Len))
                        {
                            if (poolSet.Count >= 120) break;
                            if (poolSet.Add(x.L.ObjectId)) pool.Add(x);
                        }

                        double deg = Math.PI / 180.0;
                        int totalBins = (int)Math.Round(Math.PI / (0.5 * deg));
                        int Bin(double ang)
                        {
                            double aPar = ang % Math.PI; if (aPar < 0) aPar += Math.PI;
                            return (int)Math.Round(aPar / (0.5 * deg));
                        }

                        var groups = new Dictionary<int, List<int>>();
                        for (int i = 0; i < pool.Count; i++)
                        {
                            int b = Bin(Math.Atan2(pool[i].V.Y, pool[i].V.X));
                            if (!groups.TryGetValue(b, out var list)) groups[b] = list = new List<int>();
                            list.Add(i);
                        }

                        bool SegIntersect(Point2d a0, Point2d a1, Point2d b0, Point2d b1, out Point2d ip)
                        {
                            ip = default;
                            var r = a1 - a0; var s = b1 - b0;
                            double rxs = r.X * s.Y - r.Y * s.X;
                            var qp = b0 - a0;
                            double qpxr = qp.X * r.Y - qp.Y * r.X;
                            double EPS = 1e-9;
                            if (Math.Abs(rxs) < EPS) return false;
                            double t = (qp.X * s.Y - qp.Y * s.X) / rxs;
                            double u = qpxr / rxs;
                            if (t < -1e-6 || t > 1 + 1e-6 || u < -1e-6 || u > 1 + 1e-6) return false;
                            ip = a0 + t * r;
                            return true;
                        }

                        foreach (var kvA in groups)
                        {
                            int bins90 = (int)Math.Round(90.0 / 0.5);
                            int perpKeyCenter = (kvA.Key + bins90) % totalBins;
                            var perpKeys = new int[] { (perpKeyCenter - 1 + totalBins) % totalBins, perpKeyCenter, (perpKeyCenter + 1) % totalBins };

                            if (!kvA.Value.Any()) continue;
                            var idxA = kvA.Value.OrderByDescending(i => pool[i].Len).Take(12).ToArray();

                            foreach (int keyB in perpKeys)
                            {
                                if (!groups.TryGetValue(keyB, out var idxB)) continue;
                                var idxBBest = idxB.OrderByDescending(i => pool[i].Len).Take(12).ToArray();

                                for (int i1 = 0; i1 < idxA.Length; i1++)
                                    for (int i2 = i1 + 1; i2 < idxA.Length; i2++)
                                    {
                                        var A1 = pool[idxA[i1]]; var A2 = pool[idxA[i2]];
                                        var nA = new Vector2d(-A1.V.Y, A1.V.X);
                                        if (Math.Abs(A1.V.DotProduct(A2.V)) < 0.98) continue; // Ensure lines are parallel

                                        for (int j1 = 0; j1 < idxBBest.Length; j1++)
                                            for (int j2 = j1 + 1; j2 < idxBBest.Length; j2++)
                                            {
                                                var B1 = pool[idxBBest[j1]]; var B2 = pool[idxBBest[j2]];
                                                if (Math.Abs(B1.V.DotProduct(B2.V)) < 0.98) continue; // Ensure lines are parallel

                                                if (!SegIntersect(A1.P0, A1.P1, B1.P0, B1.P1, out var C00)) continue;
                                                if (!SegIntersect(A1.P0, A1.P1, B2.P0, B2.P1, out var C01)) continue;
                                                if (!SegIntersect(A2.P0, A2.P1, B1.P0, B1.P1, out var C10)) continue;
                                                if (!SegIntersect(A2.P0, A2.P1, B2.P0, B2.P1, out var C11)) continue;

                                                var u = C01 - C00; var v = C10 - C00;
                                                double wLen = u.Length, hLen = v.Length; if (wLen < 1e-4 || hLen < 1e-4) continue;
                                                double dotArea = Math.Abs(u.X * v.Y - u.Y * v.X);
                                                double norm = wLen * hLen; if (norm <= 0) continue;
                                                double sinTheta = dotArea / norm; if (Math.Abs(sinTheta - 1.0) > 0.02) continue;
                                                double wOpp = (C11 - C10).Length, hOpp = (C11 - C01).Length;
                                                if (Math.Abs(wLen - wOpp) > 0.015 * Math.Max(wLen, wOpp)) continue;
                                                if (Math.Abs(hLen - hOpp) > 0.015 * Math.Max(hLen, hOpp)) continue;

                                                var ext = new Extents3d();
                                                ext.AddPoint(new Point3d(C00.X, C00.Y, 0));
                                                ext.AddPoint(new Point3d(C01.X, C01.Y, 0));
                                                ext.AddPoint(new Point3d(C10.X, C10.Y, 0));
                                                ext.AddPoint(new Point3d(C11.X, C11.Y, 0));

                                                double angle = Math.Atan2(u.Y, u.X);
                                                double ratio = Math.Max(wLen, hLen) / Math.Min(wLen, hLen);
                                                double score = 0.0; if (IsSheetRatio(ratio)) score += 3.0;
                                                double area = wLen * hLen; if (area > 1) score += Math.Log10(area);
                                                int hintHits = (new[] { A1, A2, B1, B2 }).Count(l => layerTokens.Any(t => l.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0));
                                                score += 0.2 * hintHits;
                                                if (attDefs.Count > 0)
                                                {
                                                    int inside = attDefs.Count(a => PointInExtents2D(a.Pos, ext));
                                                    score += Math.Min(2.0, inside * 0.3);
                                                    int tagHits = attDefs.Count(a => PointInExtents2D(a.Pos, ext) && attrTags.Contains(a.Tag));
                                                    score += Math.Min(2.0, tagHits * 0.5);
                                                }
                                                if (score >= 3.0)
                                                {
                                                    var poly = new[]
                                                    {
                                                        new Point3d(C00.X, C00.Y, 0), new Point3d(C01.X, C01.Y, 0),
                                                        new Point3d(C11.X, C11.Y, 0), new Point3d(C10.X, C10.Y, 0)
                                                    };
                                                    var boundary = new[] { A1.L.ObjectId, A2.L.ObjectId, B1.L.ObjectId, B2.L.ObjectId };
                                                    candidates.Add((ext, score, angle, poly, boundary));
                                                }
                                            }
                                    }
                            }
                        }
                    }

                    // 3) Fallback to ATTDEF cluster extents
                    if (candidates.Count == 0 && attDefs.Count >= 3)
                    {
                        var extentsBuilder = new Extents3d();
                        attDefs.ForEach(a => extentsBuilder.AddPoint(a.Pos));
                        var ext = extentsBuilder;

                        double minWidth = 10.0; // Minimum reasonable width
                        double minHeight = 7.5; // Minimum reasonable height
                        double dx = Math.Max((ext.MaxPoint.X - ext.MinPoint.X) * 0.25, minWidth);
                        double dy = Math.Max((ext.MaxPoint.Y - ext.MinPoint.Y) * 0.25, minHeight);
                        ext = new Extents3d(new Point3d(ext.MinPoint.X - dx, ext.MinPoint.Y - dy, 0), new Point3d(ext.MaxPoint.X + dx, ext.MaxPoint.Y + dy, 0));
                        var poly = new[]
                        {
                            new Point3d(ext.MinPoint.X, ext.MinPoint.Y, 0),
                            new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0),
                            new Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, 0),
                            new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0)
                        };
                        candidates.Add((ext, 2.9, 0.0, poly, Array.Empty<ObjectId>()));
                    }

                    if (candidates.Count == 0)
                    {
                        ed.WriteMessage("\nNo likely title block found in Model Space.");
                        tr.Commit();
                        return;
                    }

                    var best = candidates.OrderByDescending(c => c.Score).First();

                    // *** FIX: Use ed.CurrentUserCoordinateSystem which is a Matrix3d and has .Inverse() ***
                    var ucs = ed.CurrentUserCoordinateSystem;
                    var wcs_poly = best.Poly.Select(p => p.TransformBy(ucs.Inverse())).ToArray();

                    var expandedPoly = ExpandTitleBlockPolygon(wcs_poly);
                    ZoomToTitleBlock(ed, expandedPoly ?? wcs_poly);

                    // Select objects strictly inside the found polygon
                    var polyColl = new Point3dCollection(expandedPoly ?? best.Poly);
                    var selRes = ed.SelectCrossingPolygon(polyColl);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nNothing found inside the detected titleblock region.");
                        tr.Commit();
                        return;
                    }

                    keepIds = new HashSet<ObjectId>(selRes.Value.GetObjectIds());
                    // Ensure the border entities themselves are also kept
                    foreach (var bid in best.Boundary) if (!bid.IsNull) keepIds.Add(bid);

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in KEEPONLYTITLEBLOCKMS: {ex.Message}\n{ex.StackTrace}");
                return;
            }

            if (keepIds == null || keepIds.Count == 0)
            {
                return;
            }

            try
            {
                var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                EraseEntitiesExcept(db, ed, modelSpaceId, keepIds);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prune model space: {ex.Message}");
            }
        }

        private static Extents3d ExpandExtents(Extents3d ext, double fraction)
        {
            double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
            double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);
            double dx = Math.Max(width * fraction, 1e-4);
            double dy = Math.Max(height * fraction, 1e-4);

            return new Extents3d(
                new Point3d(ext.MinPoint.X - dx, ext.MinPoint.Y - dy, ext.MinPoint.Z),
                new Point3d(ext.MaxPoint.X + dx, ext.MaxPoint.Y + dy, ext.MaxPoint.Z));
        }

        public static void ZoomToTitleBlock(Editor ed, Point3d[] poly)
        {
            if (ed == null || poly == null || poly.Length < 3) return;

            try
            {
                var ext = new Extents3d(poly[0], poly[0]);
                for (int i = 1; i < poly.Length; i++)
                {
                    ext.AddPoint(poly[i]);
                }

                double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
                double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);

                if (width < 1e-4 || height < 1e-4)
                {
                    ed.WriteMessage("\nDetected title block area is too small, skipping zoom.");
                    return;
                }

                double margin = Math.Max(width, height) * 0.05;

                Point3d pMin = new Point3d(ext.MinPoint.X - margin, ext.MinPoint.Y - margin, ext.MinPoint.Z);
                Point3d pMax = new Point3d(ext.MaxPoint.X + margin, ext.MaxPoint.Y + margin, ext.MaxPoint.Z);

                Zoom(pMin, pMax, Point3d.Origin, 1.0);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during zoom: {ex.Message}");
            }
        }

        public static void Zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            if (acDoc == null) return;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                using (ViewTableRecord acView = ed.GetCurrentView())
                {
                    Point2d pNewCentPt;
                    double dWidth, dHeight;

                    Matrix3d matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist, acView.ViewDirection, acView.Target) * matWCS2DCS;
                    matWCS2DCS = matWCS2DCS.Inverse();

                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                        dWidth = acView.Width;
                        dHeight = acView.Height;
                    }
                    else
                    {
                        var extents = new Extents3d(pMin, pMax);
                        extents.TransformBy(matWCS2DCS);
                        dWidth = extents.MaxPoint.X - extents.MinPoint.X;
                        dHeight = extents.MaxPoint.Y - extents.MinPoint.Y;
                        pNewCentPt = new Point2d((extents.MinPoint.X + extents.MaxPoint.X) / 2.0, (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0);
                    }

                    double dViewRatio = acView.Width / acView.Height;
                    if (dWidth > (dHeight * dViewRatio))
                    {
                        dHeight = dWidth / dViewRatio;
                    }

                    acView.CenterPoint = pNewCentPt;
                    acView.Width = dWidth * dFactor;
                    acView.Height = dHeight * dFactor;

                    ed.SetCurrentView(acView);
                }
                acTrans.Commit();
            }
        }

        private static Point3d[] ExpandTitleBlockPolygon(Point3d[] poly)
        {
            if (poly == null || poly.Length != 4)
            {
                return poly;
            }

            var center = new Point3d(
                (poly[0].X + poly[1].X + poly[2].X + poly[3].X) / 4.0,
                (poly[0].Y + poly[1].Y + poly[2].Y + poly[3].Y) / 4.0,
                0
            );

            // *** FIX: Changed GetDistanceTo to DistanceTo ***
            double side1 = poly[0].DistanceTo(poly[1]);
            double side2 = poly[0].DistanceTo(poly[3]);

            double margin = Math.Max(0.01, Math.Min(side1, side2) * 0.005);

            var expandedPoly = new Point3d[4];
            for (int i = 0; i < 4; i++)
            {
                var vecFromCenter = poly[i] - center;
                if (vecFromCenter.Length < 1e-6) return poly;
                expandedPoly[i] = poly[i] + vecFromCenter.GetNormal() * margin;
            }

            return expandedPoly;
        }
    }
}