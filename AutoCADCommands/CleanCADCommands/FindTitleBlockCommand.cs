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
        // Utility: Find likely title block rectangle in THIS drawing's Model Space
        [CommandMethod("FINDTITLEBLOCKMS", CommandFlags.Modal)]
        public static void FindTitleBlockInModelSpace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;
        
            // Tokens to tailor as needed for your standards
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
                    if (db.CurrentSpaceId != ms.ObjectId)
                    {
                        ed.WriteMessage("\nPlease switch to Model Space before running KEEPONLYTITLEBLOCKMS.");
                        tr.Commit();
                        return;
                    }
        
                    // Collect attribute definitions present in Model Space (common in TB source files)
                    var attDefs = new List<(Point3d Pos, string Tag)>();
                    foreach (ObjectId id in ms)
                    {
                        if (!id.IsValid) continue;
                        DBObject dbo = tr.GetObject(id, OpenMode.ForRead, false);
                        if (dbo is AttributeDefinition ad && !ad.Invisible)
                        {
                            attDefs.Add((ad.Position, ad.Tag));
                        }
                    }
        
                    // Scan for rectangular closed polylines and score them
                    var candidates = new List<(Entity Ent, Extents3d Ext, double Score, double Angle)>();
                    foreach (ObjectId id in ms)
                    {
                        if (!id.IsValid) continue;
                        Entity ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                        if (ent == null) continue;
        
                        Extents3d? ext = TryGetExtents(ent);
                        if (ext == null) continue;
        
                        double score = 0.0;
                        double angle = 0.0;
        
                        // Layer hint
                        if (layerTokens.Any(t => ent.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0))
                            score += 0.8;
        
                        // Prefer clean rectangles (closed lwpolyline with 4 verts)
                        if (ent is Polyline pl && pl.Closed && pl.NumberOfVertices == 4)
                        {
                            var (isRect, w, h, ang) = TryRectInfo(pl);
                            if (isRect)
                            {
                                angle = ang;
                                double ratio = Math.Max(w, h) / Math.Max(1e-9, Math.Min(w, h));
                                if (IsSheetRatio(ratio)) score += 3.0; // looks like a sheet
                                score += 0.7; // clean rectangle bonus
        
                                // If attribute defs are inside the extents, bump score
                                if (attDefs.Count > 0)
                                {
                                    int inside = 0;
                                    foreach (var a in attDefs)
                                    {
                                        if (PointInExtents2D(a.Pos, ext.Value)) inside++;
                                    }
                                    // modest bump based on count
                                    score += Math.Min(2.0, inside * 0.3);
                                    // bonus if tags look like TB attributes
                                    int tagHits = attDefs.Count(a => PointInExtents2D(a.Pos, ext.Value) && attrTags.Contains(a.Tag));
                                    score += Math.Min(2.0, tagHits * 0.5);
                                }
                            }
                        }
                        // Secondary: a huge axis-aligned rectangle composed of lines/polylines
                        else if (ent is Polyline pl2 && pl2.Closed)
                        {
                            // Fallback on extents and size if not a clean 4-vertex rectangle
                            var v = ext.Value.MaxPoint - ext.Value.MinPoint;
                            double w = Math.Abs(v.X), h = Math.Abs(v.Y);
                            if (w > 0 && h > 0)
                            {
                                double ratio = Math.Max(w, h) / Math.Min(w, h);
                                if (IsSheetRatio(ratio)) score += 1.2;
                            }
                        }
        
                        // Size bias (prefer larger areas to avoid logos)
                        var s = ext.Value.MaxPoint - ext.Value.MinPoint;
                        double area = Math.Abs(s.X * s.Y);
                        if (area > 1) score += Math.Log10(area);
        
                        if (score >= 2.8)
                        {
                            candidates.Add((ent, ext.Value, score, angle));
                        }
                    }
        
                    // NEW: Detect rectangles composed of four Line segments (common border style)
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
                            if (len > 1e-6)
                            {
                                lines.Add((ln, p0, p1, v / len, len, ln.Layer));
                            }
                        }
                    }
        
                    if (lines.Count > 0)
                    {
                        // Prefer likely border layers and longer segments
                        var hinted = lines.Where(l => layerTokens.Any(t => l.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();
                        var poolSet = new HashSet<ObjectId>(hinted.Select(h => h.L.ObjectId));
                        var pool = new List<(Line L, Point2d P0, Point2d P1, Vector2d V, double Len, string Layer)>(hinted);
                        foreach (var x in lines.OrderByDescending(x => x.Len))
                        {
                            if (poolSet.Count >= 120) break; // cap
                            if (poolSet.Add(x.L.ObjectId)) pool.Add(x);
                        }
        
                        // Group by direction (parallel bins ~0.5 deg)
                        double deg = Math.PI / 180.0;
                        int totalBins = (int)Math.Round(Math.PI / (0.5 * deg)); // ~360
                        int Bin(double ang)
                        {
                            // parallel angle in [0,pi)
                            double a = ang;
                            double aPar = a % Math.PI; if (aPar < 0) aPar += Math.PI;
                            return (int)Math.Round(aPar / (0.5 * deg));
                        }
        
                        var groups = new Dictionary<int, List<int>>();
                        var angles = new List<double>();
                        for (int i = 0; i < pool.Count; i++)
                        {
                            double a = Math.Atan2(pool[i].V.Y, pool[i].V.X);
                            angles.Add(a);
                            int b = Bin(a);
                            if (!groups.TryGetValue(b, out var list)) groups[b] = list = new List<int>();
                            list.Add(i);
                        }
        
                        // helper for segment intersection
                        bool SegIntersect(Point2d a0, Point2d a1, Point2d b0, Point2d b1, out Point2d ip)
                        {
                            ip = default;
                            var r = a1 - a0; var s = b1 - b0;
                            double rxs = r.X * s.Y - r.Y * s.X;
                            double qpxr = (b0.X - a0.X) * r.Y - (b0.Y - a0.Y) * r.X;
                            double EPS = 1e-7;
                            if (Math.Abs(rxs) < EPS) return false; // parallel or colinear; skip
                            double t = ((b0.X - a0.X) * s.Y - (b0.Y - a0.Y) * s.X) / rxs;
                            double u = qpxr / rxs;
                            if (t < -1e-6 || t > 1 + 1e-6 || u < -1e-6 || u > 1 + 1e-6) return false;
                            ip = new Point2d(a0.X + t * r.X, a0.Y + t * r.Y);
                            return true;
                        }
        
                        // Try rectangle assembly from two parallel lines in A and two in B (perpendicular to A)
                        foreach (var kvA in groups)
                        {
                            // find perpendicular bins around +90 deg (+/- 1 bin)
                            int bins90 = (int)Math.Round(90.0 / 0.5); // 180 bins
                            int perpKeyCenter = (kvA.Key + bins90) % totalBins;
                            if (perpKeyCenter < 0) perpKeyCenter += totalBins;
                            int k0 = perpKeyCenter - 1; if (k0 < 0) k0 += totalBins;
                            int k1 = perpKeyCenter;
                            int k2 = perpKeyCenter + 1; if (k2 >= totalBins) k2 -= totalBins;
                            var perpKeys = new int[] { k0, k1, k2 };
                            foreach (int keyB in perpKeys)
                            {
                                if (!groups.TryGetValue(keyB, out var idxB)) continue;
        
                                var idxA = kvA.Value.OrderByDescending(i => pool[i].Len).Take(12).ToArray();
                                var idxBBest = idxB.OrderByDescending(i => pool[i].Len).Take(12).ToArray();
        
                                for (int i1 = 0; i1 < idxA.Length; i1++)
                                for (int i2 = i1 + 1; i2 < idxA.Length; i2++)
                                {
                                    var A1 = pool[idxA[i1]]; var A2 = pool[idxA[i2]];
                                    // ensure distinct and reasonably spaced
                                    // Separation between two parallel lines: project delta onto unit normal of A1
                                    var sepDeltaA = A2.P0 - A1.P0;
                                    var nA = new Vector2d(-A1.V.Y, A1.V.X); // rotate 90° CCW to get normal
                                    double sepA = Math.Abs(sepDeltaA.DotProduct(nA)); // A1.V is unit; nA is unit too
                                    if (sepA < 1e-4) continue;
        
                                    for (int j1 = 0; j1 < idxBBest.Length; j1++)
                                    for (int j2 = j1 + 1; j2 < idxBBest.Length; j2++)
                                    {
                                        var B1 = pool[idxBBest[j1]]; var B2 = pool[idxBBest[j2]];
                                        var deltaB = B2.P0 - B1.P0;
                                        var nB = new Vector2d(-B1.V.Y, B1.V.X);
                                        double distB = Math.Abs(deltaB.DotProduct(nB));
                                        if (distB < 1e-4) continue;
        
                                        // Compute the four corners
                                        if (!SegIntersect(A1.P0, A1.P1, B1.P0, B1.P1, out var C00)) continue;
                                        if (!SegIntersect(A1.P0, A1.P1, B2.P0, B2.P1, out var C01)) continue;
                                        if (!SegIntersect(A2.P0, A2.P1, B1.P0, B1.P1, out var C10)) continue;
                                        if (!SegIntersect(A2.P0, A2.P1, B2.P0, B2.P1, out var C11)) continue;
        
                                        var u = C01 - C00; var v = C10 - C00;
                                        double wLen = u.Length; double hLen = v.Length;
                                        if (wLen < 1e-4 || hLen < 1e-4) continue;
                                        // Orthogonality check
                                        double dot = Math.Abs(u.X * v.Y - u.Y * v.X); // use area magnitude to avoid normalization
                                        double norm = wLen * hLen;
                                        if (norm <= 0) continue;
                                        double sinTheta = dot / norm; // ~sin between u and v
                                        if (Math.Abs(sinTheta - 1.0) > 0.02) continue; // ~within ~1.15 degrees of 90°
        
                                        // Opposite sides roughly equal
                                        double wOpp = (C11 - C10).Length; double hOpp = (C11 - C01).Length;
                                        if (Math.Abs(wLen - wOpp) > 0.01 * Math.Max(wLen, wOpp)) continue;
                                        if (Math.Abs(hLen - hOpp) > 0.01 * Math.Max(hLen, hOpp)) continue;
        
                                        // Build extents from four corners
                                        double minX = new[] { C00.X, C01.X, C10.X, C11.X }.Min();
                                        double minY = new[] { C00.Y, C01.Y, C10.Y, C11.Y }.Min();
                                        double maxX = new[] { C00.X, C01.X, C10.X, C11.X }.Max();
                                        double maxY = new[] { C00.Y, C01.Y, C10.Y, C11.Y }.Max();
                                        var ext = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        
                                        // Score
                                        double angle = Math.Atan2(u.Y, u.X);
                                        double ratio = Math.Max(wLen, hLen) / Math.Min(wLen, hLen);
                                        double score = 0.0;
                                        if (IsSheetRatio(ratio)) score += 3.0;
                                        // size bias
                                        double area = wLen * hLen;
                                        if (area > 1) score += Math.Log10(area);
                                        // layer hint from contributing lines
                                        int hintHits = 0;
                                        if (layerTokens.Any(t => A1.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        if (layerTokens.Any(t => A2.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        if (layerTokens.Any(t => B1.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        if (layerTokens.Any(t => B2.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) hintHits++;
                                        score += 0.2 * hintHits;
                                        // attributes inside
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
                                            candidates.Add((null, ext, score, angle));
                                        }
                                    }
                                }
                            }
                        }
                    }
        
                    // If nothing obvious, build a candidate from ATTDEF cluster
                    if (candidates.Count == 0 && attDefs.Count >= 3)
                    {
                        double minX = double.PositiveInfinity, minY = double.PositiveInfinity;
                        double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity;
                        foreach (var a in attDefs)
                        {
                            minX = Math.Min(minX, a.Pos.X); minY = Math.Min(minY, a.Pos.Y);
                            maxX = Math.Max(maxX, a.Pos.X); maxY = Math.Max(maxY, a.Pos.Y);
                        }
                        // Expand a bit to cover border
                        double dx = (maxX - minX) * 0.2;
                        double dy = (maxY - minY) * 0.2;
                        var ext = new Extents3d(new Point3d(minX - dx, minY - dy, 0), new Point3d(maxX + dx, maxY + dy, 0));
                        candidates.Add((null, ext, 2.9, 0.0));
                    }
        
                    if (candidates.Count == 0)
                    {
                        ed.WriteMessage("\nNo likely title block found in Model Space.");
                    }
                    else
                    {
                        var best = candidates.OrderByDescending(c => c.Score).First();
                        var center = new Point3d(
                            (best.Ext.MinPoint.X + best.Ext.MaxPoint.X) / 2.0,
                            (best.Ext.MinPoint.Y + best.Ext.MaxPoint.Y) / 2.0,
                            0);
                        var sz = best.Ext.MaxPoint - best.Ext.MinPoint;
                        ed.WriteMessage($"\nLikely Title Block: Center {center}, Size {Math.Abs(sz.X):0.###} x {Math.Abs(sz.Y):0.###}, Rotation {best.Angle * 180/Math.PI:0.##}°");
                    }
        
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError while searching for title block: {ex.Message}");
            }
        }
    }
}

