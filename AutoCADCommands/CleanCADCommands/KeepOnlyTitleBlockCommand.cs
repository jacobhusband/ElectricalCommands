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
        [CommandMethod("ZOOMTOTBMS", CommandFlags.Modal)]
        public static void ZoomToTitleBlockInModelSpace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                SwitchToModelSpaceViewSafe(db, ed);
                var bestCandidate = FindBestTitleBlockInModelSpace(db, ed);

                if (bestCandidate.HasValue)
                {
                    ed.WriteMessage("\nZooming to model space title block...");
                    var extents = bestCandidate.Value.Ext;
                    ZoomToTitleBlock(ed, new[] { extents.MinPoint, extents.MaxPoint });
                }
                else
                {
                    ed.WriteMessage("\nCould not find a title block to zoom to in Model Space.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during zoom to model space title block: {ex.Message}");
            }
        }

        internal static (Extents3d Ext, double Score, double Angle, Point3d[] Poly, ObjectId[] Boundary)? FindBestTitleBlockInModelSpace(Database db, Editor ed)
        {
            // Tokens that often appear on border/title layers (kept from your original)
            string[] layerTokens = { "TITLE", "TBLK", "BORDER", "SHEET", "FRAME" };
            var attrTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "SHEET","SHEETNO","SHEET_NO","SHEETNUMBER","DWG","DWG_NO",
                "DRAWN","CHECKED","APPROVED","PROJECT","CLIENT","SCALE","DATE",
                "REV","REVISION","TITLE"
            };

            // --- local helpers -------------------------------------------------------
            static double Dot(Vector2d a, Vector2d b) => (a.X * b.X + a.Y * b.Y) / (a.Length * b.Length);
            static bool LayerTitleish(string layer, string[] tokens)
                => layer.Equals("DEFPOINTS", StringComparison.OrdinalIgnoreCase)
                   || tokens.Any(t => layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0);

            static bool TryIntersect(Line a, Line b, out Point3d p)
            {
                var pts = new Point3dCollection();
                a.IntersectWith(b, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);
                if (pts.Count > 0) { p = pts[0]; return true; }
                p = Point3d.Origin; return false;
            }
            // Cosine tolerances (angles): ~2°
            const double COS_PAR = 0.99939;   // cos(2°)
            const double COS_PERP = 0.03490;  // cos(88°) ~ perpendicular if |dot| <= this

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // Collect visible ATTDEFs for context scoring (unchanged)
                var attDefs = new List<(Point3d Pos, string Tag)>();
                foreach (ObjectId id in ms)
                {
                    if (!id.IsValid) continue;
                    DBObject dbo = tr.GetObject(id, OpenMode.ForRead, false);
                    if (dbo is AttributeDefinition ad && !ad.Invisible)
                        attDefs.Add((ad.Position, ad.Tag));
                }

                var candidates = new List<(Extents3d Ext, double Score, double Angle, Point3d[] Poly, ObjectId[] Boundary)>();

                // 1) Closed 4-vertex polylines (kept from your original)
                foreach (ObjectId id in ms)
                {
                    if (!(tr.GetObject(id, OpenMode.ForRead, false) is Entity ent)) continue;
                    var ext = TryGetExtents(ent);
                    if (ext == null) continue;

                    double score = 0.0, angle = 0.0;
                    if (LayerTitleish(ent.Layer, layerTokens)) score += 0.8;

                    if (ent is Polyline pl && pl.Closed && pl.NumberOfVertices == 4)
                    {
                        var (ok, w, h, ang) = TryRectInfo(pl);
                        if (!ok) continue;
                        angle = ang;
                        double ratio = Math.Max(w, h) / Math.Max(1e-9, Math.Min(w, h));
                        if (IsSheetRatio(ratio)) score += 3.0;
                        score += 0.7;

                        int inside = attDefs.Count(a => PointInExtents2D(a.Pos, ext.Value));
                        score += Math.Min(2.0, inside * 0.3);
                        int tagHits = attDefs.Count(a => PointInExtents2D(a.Pos, ext.Value) && attrTags.Contains(a.Tag));
                        score += Math.Min(2.0, tagHits * 0.5);

                        var pts = Enumerable.Range(0, 4).Select(i => pl.GetPoint3dAt(i)).ToArray();
                        var s = ext.Value.MaxPoint - ext.Value.MinPoint;
                        double area = Math.Abs(s.X * s.Y);
                        if (area > 1) score += Math.Log10(area);
                        candidates.Add((ext.Value, score, angle, pts, new[] { ent.ObjectId }));
                    }
                }

                // 2) Rectangles made of four Lines (NEW robust detector; supports DEFPOINTS & rotation)
                var lines = new List<(ObjectId Id, Line L, Point2d P0, Point2d P1, Vector2d V, double Len, string Layer)>();
                foreach (ObjectId id in ms)
                {
                    if (tr.GetObject(id, OpenMode.ForRead, false) is Line ln)
                    {
                        var p0 = new Point2d(ln.StartPoint.X, ln.StartPoint.Y);
                        var p1 = new Point2d(ln.EndPoint.X, ln.EndPoint.Y);
                        var vRaw = p1 - p0;
                        double len = vRaw.Length;
                        if (len > 1e-6)
                        {
                            var v = vRaw / len;
                            lines.Add((id, ln, p0, p1, v, len, ln.Layer));
                        }
                    }
                }

                if (lines.Count >= 4)
                {
                    // Prevent duplicates via a normalized key of the 4 ObjectIds
                    var seen = new HashSet<string>();

                    // Heuristics: favor longer border-like lines
                    var ordered = lines.OrderByDescending(x => x.Len).ToArray();

                    for (int i = 0; i < ordered.Length; i++)
                    {
                        var A = ordered[i];
                        for (int j = i + 1; j < ordered.Length; j++)
                        {
                            var B = ordered[j];
                            double dotAB = Math.Abs(Dot(A.V, B.V));

                            // A ⟂ B ?
                            if (dotAB > COS_PERP) continue;

                            // Find partner lines K ∥ A and L ∥ B
                            for (int k = 0; k < ordered.Length; k++)
                            {
                                if (k == i || k == j) continue;
                                var K = ordered[k];
                                if (Math.Abs(Math.Abs(Dot(A.V, K.V)) - 1.0) > (1.0 - COS_PAR)) continue; // not parallel to A

                                for (int l = 0; l < ordered.Length; l++)
                                {
                                    if (l == i || l == j || l == k) continue;
                                    var L = ordered[l];
                                    if (Math.Abs(Math.Abs(Dot(B.V, L.V)) - 1.0) > (1.0 - COS_PAR)) continue; // not parallel to B

                                    // Intersections: P = A∩B, Q = K∩B, R = A∩L, S = K∩L
                                    if (!TryIntersect(A.L, B.L, out var P)) continue;
                                    if (!TryIntersect(K.L, B.L, out var Q)) continue;
                                    if (!TryIntersect(A.L, L.L, out var R)) continue;
                                    if (!TryIntersect(K.L, L.L, out var S)) continue;

                                    // Basic geometry sanity: opposite corners distinct and non-degenerate
                                    double w = P.DistanceTo(R);
                                    double h = P.DistanceTo(Q);
                                    if (w < 1e-3 || h < 1e-3) continue;

                                    // Ensure the intersections actually sit within each segment (tight-ish)
                                    // (Intersect.OnBothOperands usually guarantees this, but keep a small guard)
                                    double segTol = Math.Max(0.01, 0.002 * ((A.Len + B.Len + K.Len + L.Len) / 4.0));
                                    bool OnSeg(Line ln, Point3d p) =>
                                        p.DistanceTo(ln.StartPoint) + p.DistanceTo(ln.EndPoint) <= ln.Length + segTol;

                                    if (!OnSeg(A.L, P) || !OnSeg(B.L, P)) continue;
                                    if (!OnSeg(K.L, Q) || !OnSeg(B.L, Q)) continue;
                                    if (!OnSeg(A.L, R) || !OnSeg(L.L, R)) continue;
                                    if (!OnSeg(K.L, S) || !OnSeg(L.L, S)) continue;

                                    // Build the polygon in order (clockwise-ish): P -> R -> S -> Q
                                    var poly = new[] { P, R, S, Q };

                                    // Compute extents
                                    var ext = new Extents3d(poly[0], poly[0]);
                                    for (int t = 1; t < 4; t++) ext.AddPoint(poly[t]);

                                    // Scoring
                                    double score = 2.5; // base for 4-line rectangle
                                                        // Layer boosts (DEFPOINTS or title-ish)
                                    int tlHits = 0;
                                    if (LayerTitleish(A.Layer, layerTokens)) { score += 0.5; tlHits++; }
                                    if (LayerTitleish(B.Layer, layerTokens)) { score += 0.5; tlHits++; }
                                    if (LayerTitleish(K.Layer, layerTokens)) { score += 0.5; tlHits++; }
                                    if (LayerTitleish(L.Layer, layerTokens)) { score += 0.5; tlHits++; }
                                    if (tlHits >= 3) score += 0.3;

                                    // Aspect ratio
                                    double ratio = Math.Max(w, h) / Math.Max(1e-9, Math.Min(w, h));
                                    if (IsSheetRatio(ratio)) score += 2.0;

                                    // Strongly favor ~30x42 (Arch E1). Use ±10% band to be unit-safe (inches or metric).
                                    var dims = new[] { w, h }.OrderBy(x => x).ToArray();
                                    double shortSide = dims[0], longSide = dims[1];
                                    bool near30 = shortSide >= 27 && shortSide <= 33;
                                    bool near42 = longSide >= 37.8 && longSide <= 46.2;
                                    if (near30 && near42) score += 3.0;

                                    // Attribute context inside the border
                                    int inside = attDefs.Count(a => PointInExtents2D(a.Pos, ext));
                                    score += Math.Min(2.0, inside * 0.3);
                                    int tagHits = attDefs.Count(a => PointInExtents2D(a.Pos, ext) && attrTags.Contains(a.Tag));
                                    score += Math.Min(2.0, tagHits * 0.5);

                                    // Larger borders → small logarithmic bump
                                    var s = ext.MaxPoint - ext.MinPoint;
                                    double area = Math.Abs(s.X * s.Y);
                                    if (area > 1) score += Math.Log10(area);

                                    // Angle roughly along A's direction
                                    double angle = Math.Atan2(A.V.Y, A.V.X);

                                    // Deduplicate by id set
                                    var key = string.Join("-", new[] { A.Id.Handle.Value.ToString(), B.Id.Handle.Value.ToString(), K.Id.Handle.Value.ToString(), L.Id.Handle.Value.ToString() }.OrderBy(x => x));
                                    if (seen.Add(key))
                                    {
                                        candidates.Add((ext, score, angle, poly, new[] { A.Id, B.Id, K.Id, L.Id }));
                                    }
                                }
                            }
                        }
                    }
                }

                // 3) Fallback to ATTDEF cluster if no geometry candidate
                if (candidates.Count == 0 && attDefs.Count >= 3)
                {
                    var extentsBuilder = new Extents3d();
                    attDefs.ForEach(a => extentsBuilder.AddPoint(a.Pos));
                    var ext = extentsBuilder;

                    double dx = Math.Max((ext.MaxPoint.X - ext.MinPoint.X) * 0.25, 10.0);
                    double dy = Math.Max((ext.MaxPoint.Y - ext.MinPoint.Y) * 0.25, 7.5);
                    ext = new Extents3d(
                        new Point3d(ext.MinPoint.X - dx, ext.MinPoint.Y - dy, 0),
                        new Point3d(ext.MaxPoint.X + dx, ext.MaxPoint.Y + dy, 0)
                    );
                    var poly = new[]
                    {
                ext.MinPoint,
                new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0),
                ext.MaxPoint,
                new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0)
            };
                    candidates.Add((ext, 2.9, 0.0, poly, Array.Empty<ObjectId>()));
                }

                if (candidates.Any())
                {
                    return candidates.OrderByDescending(c => c.Score).First();
                }

                tr.Commit();
                return null;
            }
        }


        // Find the titleblock region, preselect it and everything inside, then erase everything else in Model Space
        [CommandMethod("KEEPONLYTITLEBLOCKMS", CommandFlags.Modal)]
        public static void KeepOnlyTitleBlockInModelSpace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            SwitchToModelSpaceViewSafe(db, ed);

            HashSet<ObjectId> keepIds = null;
            try
            {
                var bestCandidate = FindBestTitleBlockInModelSpace(db, ed);
                if (!bestCandidate.HasValue)
                {
                    ed.WriteMessage("\nNo likely title block found in Model Space.");
                    return;
                }

                var best = bestCandidate.Value;

                var ucs = ed.CurrentUserCoordinateSystem;
                var wcs_poly = best.Poly.Select(p => p.TransformBy(ucs.Inverse())).ToArray();

                var expandedPoly = ExpandTitleBlockPolygon(wcs_poly);
                ZoomToTitleBlock(ed, expandedPoly ?? wcs_poly);

                var polyColl = new Point3dCollection(expandedPoly ?? best.Poly);
                var selRes = ed.SelectCrossingPolygon(polyColl);
                if (selRes.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nNothing found inside the detected titleblock region.");
                    return;
                }

                keepIds = new HashSet<ObjectId>(selRes.Value.GetObjectIds());
                foreach (var bid in best.Boundary) if (!bid.IsNull) keepIds.Add(bid);
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
            if (ed == null || poly == null || poly.Length < 2) return;

            try
            {
                var ext = new Extents3d(poly[0], poly[0]);
                for (int i = 1; i < poly.Length; i++)
                {
                    ext.AddPoint(poly[i]);
                }

                double margin = Math.Max(ext.MaxPoint.X - ext.MinPoint.X, ext.MaxPoint.Y - ext.MinPoint.Y) * 0.05;

                Point3d pMin = new Point3d(ext.MinPoint.X - margin, ext.MinPoint.Y - margin, 0);
                Point3d pMax = new Point3d(ext.MaxPoint.X + margin, ext.MaxPoint.Y + margin, 0);

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
                    var matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection)
                        .PreMultiplyBy(Matrix3d.Displacement(acView.Target - Point3d.Origin))
                        .PreMultiplyBy(Matrix3d.Rotation(-acView.ViewTwist, acView.ViewDirection, acView.Target))
                        .Inverse();

                    var extents = new Extents3d(pMin, pMax);
                    extents.TransformBy(matWCS2DCS);

                    double dWidth = extents.MaxPoint.X - extents.MinPoint.X;
                    double dHeight = extents.MaxPoint.Y - extents.MinPoint.Y;
                    var pNewCentPt = new Point2d((extents.MinPoint.X + extents.MaxPoint.X) / 2.0, (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0);

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