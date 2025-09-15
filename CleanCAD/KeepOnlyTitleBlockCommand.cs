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
        // Find the titleblock region, preselect it and everything inside, then erase everything else in Model Space
        [CommandMethod("KEEPONLYTITLEBLOCKMS", CommandFlags.Modal)]
        public static void KeepOnlyTitleBlockInModelSpace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;
        
            string[] layerTokens = { "TITLE", "TBLK", "BORDER", "SHEET", "FRAME" };
            var attrTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "SHEET", "SHEETNO", "SHEET_NO", "SHEETNUMBER", "DWG", "DWG_NO",
                "DRAWN", "CHECKED", "APPROVED", "PROJECT", "CLIENT", "SCALE", "DATE",
                "REV", "REVISION", "TITLE"
            };
        
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
                    if (lines.Count > 0)
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
                            double qpxr = (b0.X - a0.X) * r.Y - (b0.Y - a0.Y) * r.X;
                            double EPS = 1e-7;
                            if (Math.Abs(rxs) < EPS) return false;
                            double t = ((b0.X - a0.X) * s.Y - (b0.Y - a0.Y) * s.X) / rxs;
                            double u = qpxr / rxs;
                            if (t < -1e-6 || t > 1 + 1e-6 || u < -1e-6 || u > 1 + 1e-6) return false;
                            ip = new Point2d(a0.X + t * r.X, a0.Y + t * r.Y);
                            return true;
                        }
        
                        foreach (var kvA in groups)
                        {
                            int bins90 = (int)Math.Round(90.0 / 0.5);
                            int perpKeyCenter = (kvA.Key + bins90) % totalBins; if (perpKeyCenter < 0) perpKeyCenter += totalBins;
                            int k0 = perpKeyCenter - 1; if (k0 < 0) k0 += totalBins;
                            int k1 = perpKeyCenter;
                            int k2 = perpKeyCenter + 1; if (k2 >= totalBins) k2 -= totalBins;
                            var perpKeys = new int[] { k0, k1, k2 };
        
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
                                    var sepDeltaA = A2.P0 - A1.P0; var nA = new Vector2d(-A1.V.Y, A1.V.X);
                                    double sepA = Math.Abs(sepDeltaA.DotProduct(nA)); if (sepA < 1e-4) continue;
        
                                    for (int j1 = 0; j1 < idxBBest.Length; j1++)
                                    for (int j2 = j1 + 1; j2 < idxBBest.Length; j2++)
                                    {
                                        var B1 = pool[idxBBest[j1]]; var B2 = pool[idxBBest[j2]];
                                        var sepDeltaB = B2.P0 - B1.P0; var nB = new Vector2d(-B1.V.Y, B1.V.X);
                                        double sepB = Math.Abs(sepDeltaB.DotProduct(nB)); if (sepB < 1e-4) continue;
        
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
                                        if (Math.Abs(wLen - wOpp) > 0.01 * Math.Max(wLen, wOpp)) continue;
                                        if (Math.Abs(hLen - hOpp) > 0.01 * Math.Max(hLen, hOpp)) continue;
        
                                        double minX = new[] { C00.X, C01.X, C10.X, C11.X }.Min();
                                        double minY = new[] { C00.Y, C01.Y, C10.Y, C11.Y }.Min();
                                        double maxX = new[] { C00.X, C01.X, C10.X, C11.X }.Max();
                                        double maxY = new[] { C00.Y, C01.Y, C10.Y, C11.Y }.Max();
                                        var ext = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        
                                        double angle = Math.Atan2(u.Y, u.X);
                                        double ratio = Math.Max(wLen, hLen) / Math.Min(wLen, hLen);
                                        double score = 0.0; if (IsSheetRatio(ratio)) score += 3.0;
                                        double area = wLen * hLen; if (area > 1) score += Math.Log10(area);
                                        int hintHits = 0;
                                        if (layerTokens.Any(t => A1.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        if (layerTokens.Any(t => A2.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        if (layerTokens.Any(t => B1.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        if (layerTokens.Any(t => B2.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        score += 0.2 * hintHits;
                                        if (attDefs.Count > 0)
                                        {
                                            int inside = 0;
                                            foreach (var a in attDefs) if (PointInExtents2D(a.Pos, ext)) inside++;
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
                        double minX = double.PositiveInfinity, minY = double.PositiveInfinity;
                        double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity;
                        foreach (var a in attDefs)
                        {
                            minX = Math.Min(minX, a.Pos.X); minY = Math.Min(minY, a.Pos.Y);
                            maxX = Math.Max(maxX, a.Pos.X); maxY = Math.Max(maxY, a.Pos.Y);
                        }
                        double dx = (maxX - minX) * 0.2; double dy = (maxY - minY) * 0.2;
                        var ext = new Extents3d(new Point3d(minX - dx, minY - dy, 0), new Point3d(maxX + dx, maxY + dy, 0));
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
        
                    // Select objects strictly inside the found polygon
                    var polyColl = new Point3dCollection(best.Poly);
                    var selRes = ed.SelectWindowPolygon(polyColl);
                    if (selRes.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nNothing found inside the detected titleblock region.");
                        tr.Commit();
                        return;
                    }
        
                    var keepIds = new HashSet<ObjectId>(selRes.Value.GetObjectIds());
                    // Ensure the border entities themselves are also kept
                    foreach (var bid in best.Boundary) if (!bid.IsNull) keepIds.Add(bid);
        
                    // Pre-select and call EraseOther to remove everything else from Model Space
                    ed.SetImpliedSelection(keepIds.ToArray());
                    doc.SendStringToExecute("EraseOther ", true, false, false);
        
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in KEEPONLYTITLEBLOCKMS: {ex.Message}");
            }
        }
    }
}

