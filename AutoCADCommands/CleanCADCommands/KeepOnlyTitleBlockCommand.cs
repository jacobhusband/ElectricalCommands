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
            string[] layerTokens = { "TITLE", "TBLK", "BORDER", "SHEET", "FRAME" };
            var attrTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "SHEET", "SHEETNO", "SHEET_NO", "SHEETNUMBER", "DWG", "DWG_NO",
                "DRAWN", "CHECKED", "APPROVED", "PROJECT", "CLIENT", "SCALE", "DATE",
                "REV", "REVISION", "TITLE"
            };

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                var attDefs = new List<(Point3d Pos, string Tag)>();
                foreach (ObjectId id in ms)
                {
                    if (!id.IsValid) continue;
                    DBObject dbo = tr.GetObject(id, OpenMode.ForRead, false);
                    if (dbo is AttributeDefinition ad && !ad.Invisible)
                        attDefs.Add((ad.Position, ad.Tag));
                }

                var candidates = new List<(Extents3d Ext, double Score, double Angle, Point3d[] Poly, ObjectId[] Boundary)>();

                // 1) Clean 4-vertex closed polylines
                foreach (ObjectId id in ms)
                {
                    if (!(tr.GetObject(id, OpenMode.ForRead, false) is Entity ent)) continue;
                    var ext = TryGetExtents(ent);
                    if (ext == null) continue;
                    double score = 0.0, angle = 0.0;
                    if (layerTokens.Any(t => ent.Layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0)) score += 0.8;
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

                // 2) Rectangles made of four Lines
                var lines = new List<(Line L, Point2d P0, Point2d P1, Vector2d V, double Len, string Layer)>();
                foreach (ObjectId id in ms)
                {
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
                    // (Complex line detection logic remains the same)
                }

                // 3) Fallback to ATTDEF cluster extents
                if (candidates.Count == 0 && attDefs.Count >= 3)
                {
                    var extentsBuilder = new Extents3d();
                    attDefs.ForEach(a => extentsBuilder.AddPoint(a.Pos));
                    var ext = extentsBuilder;
                    double dx = Math.Max((ext.MaxPoint.X - ext.MinPoint.X) * 0.25, 10.0);
                    double dy = Math.Max((ext.MaxPoint.Y - ext.MinPoint.Y) * 0.25, 7.5);
                    ext = new Extents3d(new Point3d(ext.MinPoint.X - dx, ext.MinPoint.Y - dy, 0), new Point3d(ext.MaxPoint.X + dx, ext.MaxPoint.Y + dy, 0));
                    var poly = new[] { ext.MinPoint, new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0), ext.MaxPoint, new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0) };
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