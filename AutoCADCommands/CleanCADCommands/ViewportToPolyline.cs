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
                    var regions = new List<ViewportRegion>();
                    var msId = SymbolUtilityServices.GetBlockModelSpaceId(db);

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        ObjectId debugLayerId = EnsureDebugLayer(db, tr);
                        var msBtr = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForWrite);

                        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                        foreach (DBDictionaryEntry kv in layoutDict)
                        {
                            var layout = (Layout)tr.GetObject(kv.Value, OpenMode.ForRead);
                            if (layout == null) continue;
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) continue;

                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

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

                            foreach (var vpId in vpIds)
                            {
                                var vp = (Viewport)tr.GetObject(vpId, OpenMode.ForRead);
                                if (!vp.On) continue;
                                if (vp.Width <= 0 || vp.Height <= 0) continue;

                                var psPoints = GetViewportBoundaryPointsInPaper(vp, tr);
                                if (psPoints == null || psPoints.Count < 3) continue;

                                // Build the PS->MS transform using the Arbitrary Axis Algorithm
                                // (AAA) — the same algorithm AutoCAD uses internally for
                                // CHSPACE and view transformations. The original
                                // PaperPolygonToModel_Correct used a different threshold and
                                // GetPerpendicularVector(), which produced wrong polygons for
                                // viewports whose ViewDirection wasn't exactly (0,0,1).
                                Matrix3d msFromPs;
                                try
                                {
                                    msFromPs = ComputeMsFromPsMatrixAaa(vp);
                                }
                                catch (System.Exception mex)
                                {
                                    ed.WriteMessage($"\nVP2PL: PS->MS matrix build failed for viewport on '{layout.LayoutName}': {mex.Message}");
                                    continue;
                                }

                                // Apply the transform to every PS boundary point.
                                var msVerts = new List<Point3d>();
                                Point3d? lastPt = null;
                                foreach (var ps in psPoints)
                                {
                                    var msPt = new Point3d(ps.X, ps.Y, 0).TransformBy(msFromPs);
                                    if (lastPt.HasValue && lastPt.Value.DistanceTo(msPt) < 1e-9) continue;
                                    msVerts.Add(msPt);
                                    lastPt = msPt;
                                }
                                if (msVerts.Count < 3) continue;

                                var msVertsArr = msVerts.ToArray();

                                // Drop a temporary Polyline in modelspace so the selection
                                // region participates in extents before the screen-bound CP select.
                                ObjectId plId = CreatePolylineInModelspace(tr, msBtr, msVertsArr, debugLayerId);
                                if (plId.IsNull) continue;

                                // Also build the Point3dCollection used by SelectCrossingPolygon.
                                var coll = new Point3dCollection();
                                foreach (var p in msVertsArr) coll.Add(p);

                                regions.Add(new ViewportRegion(layout.LayoutName, vp.ObjectId, coll, plId));
                                ed.WriteMessage($"\nVP2PL: Created modelspace region for layout '{layout.LayoutName}' viewport {vp.Number} with {msVerts.Count} vertices.");
                            }
                        }

                        tr.Commit();
                    }

                    if (regions.Count == 0)
                    {
                        ed.WriteMessage("\nVP2PL: No eligible viewport regions found on any paper space layout. Erasing all objects in Model Space.");
                        int erasedAll = EraseEntitiesExcept(db, msId, new HashSet<ObjectId>());
                        if (erasedAll > 0)
                            ed.WriteMessage($"\nVP2PL: Erased all {erasedAll} object(s) from Model Space.");
                        else
                            ed.WriteMessage("\nVP2PL: Model Space was already empty.");
                        return;
                    }

                    if (!SwitchToTrueModelSpaceView(db, ed))
                    {
                        ed.WriteMessage("\nVP2PL: Could not activate the Model tab. Aborting before modelspace cleanup.");
                        EraseViewportRegionHelpers(db, regions);
                        return;
                    }

                    // SelectCrossingPolygon is screen-bound, so make all modelspace geometry
                    // visible after switching out of any paper-space/floating viewport context.
                    ZoomToModelExtents(db, ed);

                    ed.WriteMessage($"\nVP2PL: Selecting modelspace entities visible in {regions.Count} viewport region(s)...");

                    var keepIds = new HashSet<ObjectId>();
                    var helperPolyIds = new HashSet<ObjectId>(regions.Select(r => r.HelperPolylineId));
                    int selectedEntityCount = SelectViewportRegionContents(ed, regions, helperPolyIds, keepIds);

                    if (selectedEntityCount == 0)
                    {
                        ed.WriteMessage("\nVP2PL: Nothing crossed any viewport polygon. Erasing all objects in Model Space.");
                        int erasedAll = EraseEntitiesExcept(db, msId, new HashSet<ObjectId>());
                        ed.WriteMessage($"\nVP2PL: Erased all {erasedAll} object(s) from Model Space.");
                        return;
                    }

                    AddProtectedModelSpaceEntities(db, msId, keepIds);

                    int erasedCount = EraseEntitiesExcept(db, msId, keepIds);
                    ed.WriteMessage($"\nVP2PL: Used and removed {helperPolyIds.Count} temporary viewport helper polyline(s); kept {selectedEntityCount} entit(ies) inside any viewport region; erased {erasedCount} others.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nVP2PL failed: {ex.Message}");
            }
        }

        // -------------------- Updated / New Helpers --------------------

        private sealed class ViewportRegion
        {
            internal ViewportRegion(string layoutName, ObjectId viewportId, Point3dCollection polygon, ObjectId helperPolylineId)
            {
                LayoutName = layoutName ?? string.Empty;
                ViewportId = viewportId;
                Polygon = polygon;
                HelperPolylineId = helperPolylineId;
            }

            internal string LayoutName { get; }
            internal ObjectId ViewportId { get; }
            internal Point3dCollection Polygon { get; }
            internal ObjectId HelperPolylineId { get; }
        }

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
                    if (!id.IsValid || keep.Contains(id) || SimplerCommands.IsProtectedEmbeddedOle(id)) continue;

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

        private static bool SwitchToTrueModelSpaceView(Database db, Editor ed)
        {
            try
            {
                var modelId = SymbolUtilityServices.GetBlockModelSpaceId(db);

                try
                {
                    if (Convert.ToInt16(Application.GetSystemVariable("TILEMODE")) == 0 &&
                        Convert.ToInt16(Application.GetSystemVariable("CVPORT")) != 1)
                    {
                        ed.SwitchToPaperSpace();
                    }
                }
                catch { }

                try
                {
                    if (Convert.ToInt16(Application.GetSystemVariable("TILEMODE")) == 0)
                        Application.SetSystemVariable("TILEMODE", 1);
                }
                catch { }

                if (db.CurrentSpaceId != modelId)
                {
                    try { ed.SwitchToModelSpace(); } catch { }
                }

                try { ed.Regen(); } catch { }

                bool inModelTab = true;
                try { inModelTab = Convert.ToInt16(Application.GetSystemVariable("TILEMODE")) == 1; } catch { }

                return inModelTab && db.CurrentSpaceId == modelId;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nVP2PL: Modelspace activation failed: {ex.Message}");
                return false;
            }
        }

        private static void SwitchToModelSpaceViewSafe(Database db, Editor ed)
        {
            SwitchToTrueModelSpaceView(db, ed);
        }

        private static int SelectViewportRegionContents(
            Editor ed,
            IEnumerable<ViewportRegion> regions,
            HashSet<ObjectId> helperPolylineIds,
            HashSet<ObjectId> keepIds)
        {
            int selectedEntityCount = 0;
            var originalUcs = ed.CurrentUserCoordinateSystem;

            try
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;

                int regionNumber = 0;
                foreach (var region in regions)
                {
                    regionNumber++;
                    PromptSelectionResult selRes;
                    try
                    {
                        selRes = ed.SelectCrossingPolygon(region.Polygon);
                    }
                    catch (System.Exception exSel)
                    {
                        ed.WriteMessage($"\nVP2PL: SelectCrossingPolygon failed for layout '{region.LayoutName}' region {regionNumber}: {exSel.Message}");
                        continue;
                    }

                    if (selRes.Status != PromptStatus.OK || selRes.Value == null)
                    {
                        ed.WriteMessage($"\nVP2PL: Layout '{region.LayoutName}' region {regionNumber} returned no entities (status={selRes.Status}).");
                        continue;
                    }

                    var ids = selRes.Value.GetObjectIds();
                    int added = 0;
                    foreach (var id in ids)
                    {
                        if (!id.IsValid || id.IsErased) continue;
                        if (helperPolylineIds != null && helperPolylineIds.Contains(id)) continue;

                        if (keepIds.Add(id))
                        {
                            added++;
                            selectedEntityCount++;
                        }
                    }
                    ed.WriteMessage($"\nVP2PL: Layout '{region.LayoutName}' region {regionNumber} selected {ids.Length} entit(ies), kept {added} new non-helper entit(ies).");
                }
            }
            finally
            {
                ed.CurrentUserCoordinateSystem = originalUcs;
            }

            return selectedEntityCount;
        }

        private static void AddProtectedModelSpaceEntities(Database db, ObjectId msId, HashSet<ObjectId> keepIds)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForRead);
                foreach (ObjectId entId in ms)
                {
                    if (!entId.IsValid) continue;
                    if (SimplerCommands.IsProtectedEmbeddedOle(entId))
                        keepIds.Add(entId);
                }
                tr.Commit();
            }
        }

        private static int EraseViewportRegionHelpers(Database db, IEnumerable<ViewportRegion> regions)
        {
            var helperIds = regions?
                .Select(r => r.HelperPolylineId)
                .Where(id => id.IsValid && !id.IsErased)
                .ToList();

            if (helperIds == null || helperIds.Count == 0)
                return 0;

            int erased = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layersToRelock = new HashSet<ObjectId>();

                foreach (var id in helperIds)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased) continue;

                    var layer = tr.GetObject(ent.LayerId, OpenMode.ForRead, false) as LayerTableRecord;
                    if (layer != null && layer.IsLocked)
                    {
                        layer.UpgradeOpen();
                        layer.IsLocked = false;
                        layersToRelock.Add(layer.ObjectId);
                    }

                    ent.UpgradeOpen();
                    ent.Erase();
                    erased++;
                }

                foreach (var layerId in layersToRelock)
                {
                    var layer = tr.GetObject(layerId, OpenMode.ForWrite, false) as LayerTableRecord;
                    if (layer != null)
                        layer.IsLocked = true;
                }

                tr.Commit();
            }

            return erased;
        }

        // Build the paper-space -> world-space transformation matrix for a viewport using
        // AutoCAD's Arbitrary Axis Algorithm (AAA) for the view basis. This is the same
        // algorithm AutoCAD's CHSPACE command uses internally.
        //
        // The transformation chain is:
        //   1. Translate PS coords by -CenterPoint  (origin at viewport center, in PS units)
        //   2. Scale by ViewHeight/Height          (PS units -> DCS units)
        //   3. Translate by +ViewCenter            (DCS origin at view-center)
        //   4. Rotate by -TwistAngle around the view's Z axis (apply view twist)
        //   5. Rotate (u,v,0) into the world's view basis (DCS -> WCS rotation)
        //   6. Translate by +ViewTarget            (WCS position of the view target)
        //
        // The AAA picks the view's X axis stably regardless of how close ViewDirection
        // is to ±Z, avoiding the GetPerpendicularVector() ambiguity that produced wrong
        // polygons for the E02.00 viewport in earlier iterations.
        private static Matrix3d ComputeMsFromPsMatrixAaa(Viewport vp)
        {
            const double aaaTol = 1.0 / 64.0;

            Vector3d az = vp.ViewDirection.GetNormal();

            // AAA: pick the view X axis based on whether ViewDirection is "near vertical".
            Vector3d ax = (Math.Abs(az.X) < aaaTol && Math.Abs(az.Y) < aaaTol)
                ? Vector3d.YAxis.CrossProduct(az).GetNormal()
                : Vector3d.ZAxis.CrossProduct(az).GetNormal();
            Vector3d ay = az.CrossProduct(ax).GetNormal();

            // Apply twist rotation around az.
            double t = vp.TwistAngle;
            double c = Math.Cos(t), s = Math.Sin(t);
            Vector3d xView = ax * c + ay * s;
            Vector3d yView = ay * c - ax * s;

            double muPerPu = vp.ViewHeight / vp.Height;

            Point3d target = vp.ViewTarget;
            Point2d vc = vp.ViewCenter;
            Point3d cp = vp.CenterPoint;

            // Compose the matrices in order. Note Matrix3d.Multiply applies right-to-left:
            // result * point applies T1 first, then T2, etc. So the composition is
            // T_target * R_dcs_to_wcs * T_vc * S * T_minus_cp.
            Matrix3d t1 = Matrix3d.Displacement(new Vector3d(-cp.X, -cp.Y, -cp.Z));
            Matrix3d sc = Matrix3d.Scaling(muPerPu, Point3d.Origin);
            Matrix3d t2 = Matrix3d.Displacement(new Vector3d(vc.X, vc.Y, 0));

            // Build the DCS -> WCS rotation. It maps (1,0,0) -> xView, (0,1,0) -> yView,
            // (0,0,1) -> az. AlignCoordinateSystem does exactly this.
            Matrix3d rot = Matrix3d.AlignCoordinateSystem(
                Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                Point3d.Origin, xView, yView, az);

            Matrix3d t3 = Matrix3d.Displacement(target.GetAsVector());

            return t3 * rot * t2 * sc * t1;
        }

        // Ensure a 'VP2PL-DEBUG' layer exists. Creates it with ACI cyan (4) if missing.
        // Returns the layer's ObjectId so polylines can be assigned LayerId directly.
        private static ObjectId EnsureDebugLayer(Database db, Transaction tr)
        {
            const string layerName = "VP2PL-DEBUG";
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(layerName))
                return lt[layerName];

            lt.UpgradeOpen();
            var ltr = new LayerTableRecord
            {
                Name = layerName,
                Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                    Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 4)
            };
            var id = lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
            return id;
        }

        // Build a closed Polyline from the given vertices and append it to the given
        // BlockTableRecord (typically modelspace) on the given layer. Z values on the
        // vertices are dropped — Polyline is 2D at elevation 0.
        private static ObjectId CreatePolylineInModelspace(
            Transaction tr,
            BlockTableRecord msBtr,
            Point3d[] verts,
            ObjectId layerId)
        {
            if (verts == null || verts.Length < 3) return ObjectId.Null;

            var pl = new Polyline(verts.Length);
            for (int i = 0; i < verts.Length; i++)
            {
                pl.AddVertexAt(i, new Point2d(verts[i].X, verts[i].Y), 0, 0, 0);
            }
            pl.Closed = true;
            if (!layerId.IsNull) pl.LayerId = layerId;

            var id = msBtr.AppendEntity(pl);
            tr.AddNewlyCreatedDBObject(pl, true);
            return id;
        }

        // Zoom the active view to db.Extmin/Extmax + 5% margin and regen so the change
        // takes effect immediately. This is necessary because Editor.SelectCrossingPolygon
        // only returns entities whose display geometry lies within the active view's
        // screen extents (acedSSGet with CP mode is screen-bound).
        private static void ZoomToModelExtents(Database db, Editor ed)
        {
            try
            {
                // Make sure the cached extents reflect the helper polylines we just added.
                try { ed.Regen(); } catch { }

                var dbMin = db.Extmin;
                var dbMax = db.Extmax;
                double width = dbMax.X - dbMin.X;
                double height = dbMax.Y - dbMin.Y;
                if (width <= 1e-9 || height <= 1e-9) return;

                double mx = width * 0.05;
                double my = height * 0.05;
                var center = new Point2d((dbMin.X + dbMax.X) * 0.5, (dbMin.Y + dbMax.Y) * 0.5);

                using (var view = ed.GetCurrentView())
                {
                    view.CenterPoint = center;
                    view.Width = width + 2.0 * mx;
                    view.Height = height + 2.0 * my;
                    ed.SetCurrentView(view);
                }
                try { ed.Regen(); } catch { }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nVP2PL: Zoom to extents failed: {ex.Message}");
            }
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

            pl.Closed = p2.Closed || verts[0].pt.GetDistanceTo(verts[^1].pt) <= 1e-9;
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
                    if (chain[0].s.DistanceTo(chain[^1].e) < 1e-6) pl.Closed = true;
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
                var last = chain[^1];
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
