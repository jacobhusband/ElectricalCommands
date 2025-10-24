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

        private static void ExplodeAllBlockReferences(Database db, Editor ed)
        {
            int pass = 0;
            const int maxPasses = 10; // Safety break for deep nesting
            bool wasSomethingExploded;
            var processedHandles = new HashSet<long>();

            ed.WriteMessage("\nAnalyzing and exploding block references for title block detection...");

            do
            {
                wasSomethingExploded = false;
                pass++;

                var blockRefIds = new List<ObjectId>();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var msId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                    var ms = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForRead);

                    foreach (ObjectId id in ms)
                    {
                        if (id.ObjectClass.DxfName == "INSERT" && !processedHandles.Contains(id.Handle.Value))
                        {
                            blockRefIds.Add(id);
                        }
                    }
                    tr.Commit();
                }

                if (blockRefIds.Count == 0)
                {
                    break;
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var msId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                    var ms = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForWrite);

                    foreach (var id in blockRefIds)
                    {
                        processedHandles.Add(id.Handle.Value);
                        if (id.IsErased) continue;

                        try
                        {
                            var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            if (btr.IsFromExternalReference || btr.IsLayout) continue;

                            var explodedObjects = new DBObjectCollection();
                            br.Explode(explodedObjects);

                            foreach (DBObject obj in explodedObjects)
                            {
                                var ent = obj as Entity;
                                if (ent != null)
                                {
                                    ent.Layer = br.Layer; // Inherit layer from block reference
                                    ms.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                }
                                else
                                {
                                    obj.Dispose();
                                }
                            }

                            br.UpgradeOpen();
                            br.Erase();
                            wasSomethingExploded = true;
                        }
                        catch
                        {
                            // Ignore errors (e.g., non-explodable blocks like MInsert)
                        }
                    }
                    tr.Commit();
                }
            } while (wasSomethingExploded && pass < maxPasses);
        }


        internal static (Extents3d Ext, double Score, double Angle, Point3d[] Poly, ObjectId[] Boundary)? FindBestTitleBlockInModelSpace(Database db, Editor ed)
        {
            // First, explode all block references to get primitive geometry.
            ExplodeAllBlockReferences(db, ed);

            string[] layerTokens = { "TITLE", "TBLK", "BORDER", "SHEET", "FRAME", "TB" };
            var attrTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "SHEET","SHEETNO","SHEET_NO","SHEETNUMBER","DWG","DWG_NO",
                "DRAWN","CHECKED","APPROVED","PROJECT","CLIENT","SCALE","DATE",
                "REV","REVISION","TITLE"
            };

            static bool LayerTitleish(string layer, string[] tokens)
                => layer.Equals("DEFPOINTS", StringComparison.OrdinalIgnoreCase)
                   || tokens.Any(t => layer.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                var attDefs = new List<(Point3d Pos, string Tag)>();
                foreach (ObjectId id in ms)
                {
                    if (id.IsValid && tr.GetObject(id, OpenMode.ForRead, false) is AttributeDefinition ad && !ad.Invisible)
                        attDefs.Add((ad.Position, ad.Tag));
                }

                var candidates = new List<(Extents3d Ext, double Score, double Angle, Point3d[] Poly, ObjectId[] Boundary)>();

                // --- STRATEGY 1: Closed 4-vertex polylines (Fast and reliable for ideal cases) ---
                foreach (ObjectId id in ms)
                {
                    if (!(tr.GetObject(id, OpenMode.ForRead, false) is Polyline pl)) continue;
                    if (!pl.Closed || pl.NumberOfVertices != 4) continue;

                    var ext = TryGetExtents(pl);
                    if (ext == null) continue;

                    var (ok, w, h, ang) = TryRectInfo(pl);
                    if (!ok) continue;

                    double score = 0.0;
                    if (LayerTitleish(pl.Layer, layerTokens)) score += 0.8;
                    double ratio = Math.Max(w, h) / Math.Max(1e-9, Math.Min(w, h));
                    if (IsSheetRatio(ratio)) score += 3.0;
                    score += 0.7;

                    int inside = attDefs.Count(a => PointInExtents2D(a.Pos, ext.Value));
                    score += Math.Min(2.0, inside * 0.3);
                    int tagHits = attDefs.Count(a => PointInExtents2D(a.Pos, ext.Value) && attrTags.Contains(a.Tag));
                    score += Math.Min(2.0, tagHits * 0.5);

                    var pts = Enumerable.Range(0, 4).Select(i => pl.GetPoint3dAt(i)).ToArray();
                    candidates.Add((ext.Value, score, ang, pts, new[] { pl.ObjectId }));
                }

                // --- STRATEGY 2: Find boundary from outermost horizontal/vertical lines (NEW & ROBUST) ---
                var allLines = new List<(Point3d P1, Point3d P2, ObjectId Id, string Layer)>();
                var allEntitiesExtents = new Extents3d();
                bool hasGeometry = false;

                foreach (ObjectId id in ms)
                {
                    if (!(tr.GetObject(id, OpenMode.ForRead, false) is Entity ent)) continue;
                    var ext = TryGetExtents(ent);
                    if (ext == null) continue;
                    var size = ext.Value.MaxPoint - ext.Value.MinPoint;
                    if (Math.Abs(size.X) < 1e-6 && Math.Abs(size.Y) < 1e-6) continue;

                    allEntitiesExtents.AddExtents(ext.Value);
                    hasGeometry = true;

                    if (ent is Line line)
                    {
                        allLines.Add((line.StartPoint, line.EndPoint, line.ObjectId, line.Layer));
                    }
                    else if (ent is Polyline pl)
                    {
                        if (pl.NumberOfVertices < 2) continue;
                        for (int i = 0; i < pl.NumberOfVertices; i++)
                        {
                            if (pl.GetSegmentType(i) == SegmentType.Line)
                            {
                                var seg = pl.GetLineSegmentAt(i);
                                allLines.Add((seg.StartPoint, seg.EndPoint, pl.ObjectId, pl.Layer));
                            }
                        }
                    }
                }

                if (hasGeometry)
                {
                    double totalWidth = allEntitiesExtents.MaxPoint.X - allEntitiesExtents.MinPoint.X;
                    double totalHeight = allEntitiesExtents.MaxPoint.Y - allEntitiesExtents.MinPoint.Y;
                    double orthoTol = 0.01;
                    double edgeTol = Math.Min(totalWidth, totalHeight) * 0.025;

                    var horizontalLines = allLines.Where(l => Math.Abs(l.P1.Y - l.P2.Y) < orthoTol).ToList();
                    var verticalLines = allLines.Where(l => Math.Abs(l.P1.X - l.P2.X) < orthoTol).ToList();

                    var bottomLines = horizontalLines.Where(l => Math.Abs((l.P1.Y + l.P2.Y) / 2 - allEntitiesExtents.MinPoint.Y) < edgeTol).ToList();
                    var topLines = horizontalLines.Where(l => Math.Abs((l.P1.Y + l.P2.Y) / 2 - allEntitiesExtents.MaxPoint.Y) < edgeTol).ToList();
                    var leftLines = verticalLines.Where(l => Math.Abs((l.P1.X + l.P2.X) / 2 - allEntitiesExtents.MinPoint.X) < edgeTol).ToList();
                    var rightLines = verticalLines.Where(l => Math.Abs((l.P1.X + l.P2.X) / 2 - allEntitiesExtents.MaxPoint.X) < edgeTol).ToList();

                    if (bottomLines.Any() && topLines.Any() && leftLines.Any() && rightLines.Any())
                    {
                        double x_min = leftLines.Average(l => (l.P1.X + l.P2.X) / 2);
                        double x_max = rightLines.Average(l => (l.P1.X + l.P2.X) / 2);
                        double y_min = bottomLines.Average(l => (l.P1.Y + l.P2.Y) / 2);
                        double y_max = topLines.Average(l => (l.P1.Y + l.P2.Y) / 2);

                        var poly = new[]
                        {
                            new Point3d(x_min, y_min, 0), new Point3d(x_max, y_min, 0),
                            new Point3d(x_max, y_max, 0), new Point3d(x_min, y_max, 0)
                        };
                        var ext = new Extents3d(poly[0], poly[2]);
                        var boundaryIds = bottomLines.Select(l => l.Id).Concat(topLines.Select(l => l.Id))
                                            .Concat(leftLines.Select(l => l.Id)).Concat(rightLines.Select(l => l.Id))
                                            .Distinct().ToArray();

                        double score = 5.0;
                        double w = x_max - x_min;
                        double h = y_max - y_min;
                        if (w > 1.0 && h > 1.0)
                        {
                            double ratio = Math.Max(w, h) / Math.Min(w, h);
                            if (IsSheetRatio(ratio)) score += 3.0;

                            var dims = new[] { w, h }.OrderBy(x => x).ToArray();
                            bool near30 = dims[0] >= 27 && dims[0] <= 33;
                            bool near42 = dims[1] >= 37.8 && dims[1] <= 46.2;
                            if (near30 && near42) score += 4.0;

                            if (poly[0].DistanceTo(Point3d.Origin) < 5.0) score += 2.0;
                        }

                        int titleLayerLines = bottomLines.Concat(topLines).Concat(leftLines).Concat(rightLines)
                                              .Count(l => LayerTitleish(l.Layer, layerTokens));
                        score += Math.Min(2.0, titleLayerLines * 0.5);

                        int inside = attDefs.Count(a => PointInExtents2D(a.Pos, ext));
                        score += Math.Min(2.0, inside * 0.3);
                        int tagHits = attDefs.Count(a => PointInExtents2D(a.Pos, ext) && attrTags.Contains(a.Tag));
                        score += Math.Min(2.0, tagHits * 0.5);

                        candidates.Add((ext, score, 0.0, poly, boundaryIds));
                    }
                }

                // --- STRATEGY 3: Fallback to ATTDEF cluster if no geometry candidate ---
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

            // --- MODIFIED: Prompt user for boundary instead of auto-detection ---

            // 1. Prompt for the first corner
            var ppo = new PromptPointOptions("\nSelect the first corner of the title block boundary:");
            var ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nCommand cancelled.");
                return;
            }
            Point3d corner1 = ppr.Value;

            // 2. Prompt for the second corner with a rubber-band line
            var pco = new PromptCornerOptions("\nSelect the opposite corner of the title block boundary:", corner1)
            {
                UseDashedLine = true
            };
            var pcr = ed.GetCorner(pco);
            if (pcr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nCommand cancelled.");
                return;
            }
            Point3d corner2 = pcr.Value;

            // NEW: Define the title block polygon from the user's points and store it for later use.
            double minX = Math.Min(corner1.X, corner2.X);
            double minY = Math.Min(corner1.Y, corner2.Y);
            double maxX = Math.Max(corner1.X, corner2.X);
            double maxY = Math.Max(corner1.Y, corner2.Y);

            SimplerCommands._lastFoundTitleBlockPoly = new[] {
                new Point3d(minX, minY, 0),
                new Point3d(maxX, minY, 0),
                new Point3d(maxX, maxY, 0),
                new Point3d(minX, maxY, 0)
            };

            HashSet<ObjectId> keepIds = null;
            try
            {
                // 3. Define the selection polygon from the stored points
                var polyColl = new Point3dCollection(SimplerCommands._lastFoundTitleBlockPoly);

                // 4. Select all entities within the defined polygon
                var selRes = ed.SelectCrossingPolygon(polyColl);

                if (selRes.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nNothing was found inside the specified region.");

                    // --- CORRECTED KEYWORD PROMPT ---
                    var pko = new PromptKeywordOptions("\nErase all objects in Model Space?");
                    pko.Keywords.Add("Yes");
                    pko.Keywords.Add("No");
                    pko.Keywords.Default = "No";
                    pko.AppendKeywordsToMessage = true; // Let AutoCAD format the prompt like [Yes/No] <No>:

                    var pkr = ed.GetKeywords(pko);
                    if (pkr.Status == PromptStatus.OK && pkr.StringResult == "Yes")
                    {
                        keepIds = new HashSet<ObjectId>(); // Empty set will cause everything to be erased
                    }
                    else
                    {
                        ed.WriteMessage("\nOperation cancelled.");
                        return;
                    }
                }
                else
                {
                    keepIds = new HashSet<ObjectId>(selRes.Value.GetObjectIds());
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during selection: {ex.Message}\n{ex.StackTrace}");
                return;
            }

            if (keepIds == null)
            {
                return; // Should not be reached due to logic above, but here as a safeguard
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