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
        // Static variables to pass state between commands
        private static HashSet<ObjectId> _blockIdsBeforeBind = new HashSet<ObjectId>();
        private static HashSet<ObjectId> _originalXrefIds = new HashSet<ObjectId>();
        private static HashSet<ObjectId> _imageDefsToPurge = new HashSet<ObjectId>(); // Our definitive "kill list"
        internal static bool RunKeepOnlyAfterFinalize = false;
        internal static bool SkipBindDuringFinalize = false;
        internal static bool ForceDetachOriginalXrefs = false;


        private static Extents3d? TryGetExtents(Entity ent)
        {
            try { return ent.GeometricExtents; }
            catch { try { return ent.Bounds; } catch { return (Extents3d?)null; } }
        }

        private static bool PointInExtents2D(Point3d p, Extents3d e)
        {
            return p.X >= e.MinPoint.X - 1e-9 && p.X <= e.MaxPoint.X + 1e-9 &&
                   p.Y >= e.MinPoint.Y - 1e-9 && p.Y <= e.MaxPoint.Y + 1e-9;
        }

        private static bool IsSheetRatio(double r)
        {
            if (r <= 0) return false;
            double[] target = { 1.4142, 1.5, 1.3333, 1.2941, 1.5455 }; // ISO A, Arch D, Arch E, ANSI A/C/E, ANSI B/D
            return target.Any(t => Math.Abs(r - t) <= 0.12);
        }

        private static (bool ok, double w, double h, double angle) TryRectInfo(Polyline pl)
        {
            if (pl.NumberOfVertices != 4 || !pl.Closed) return (false, 0, 0, 0);
            var p0 = pl.GetPoint2dAt(0); var p1 = pl.GetPoint2dAt(1); var p2 = pl.GetPoint2dAt(2); var p3 = pl.GetPoint2dAt(3);
            var v0 = p1 - p0; var v1 = p2 - p1; var v2 = p3 - p2; var v3 = p0 - p3;
            double w = v0.Length; double h = v1.Length;
            if (w <= 1e-9 || h <= 1e-9) return (false, 0, 0, 0);
            bool ortho = Math.Abs(v0.X * v1.X + v0.Y * v1.Y) <= 1e-3 * w * h &&
                         Math.Abs(v1.X * v2.X + v1.Y * v2.Y) <= 1e-3 * w * h &&
                         Math.Abs(v2.X * v3.X + v2.Y * v3.Y) <= 1e-3 * w * h;
            if (!ortho) return (false, 0, 0, 0);
            var longEdge = w >= h ? v0 : v1;
            double angle = Math.Atan2(longEdge.Y, longEdge.X);
            return (true, w, h, angle);
        }

        internal static int EraseEntitiesExcept(Database db, Editor ed, ObjectId spaceId, HashSet<ObjectId> idsToKeep)
        {
            int erasedCount = 0;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var space = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                    var layersToUnlock = new HashSet<ObjectId>();
                    var idsToErase = new List<ObjectId>();

                    foreach (ObjectId id in space)
                    {
                        if (idsToKeep != null && idsToKeep.Contains(id))
                        {
                            continue;
                        }

                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        idsToErase.Add(id);
                        var layer = tr.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layer != null && layer.IsLocked)
                        {
                            layersToUnlock.Add(ent.LayerId);
                        }
                    }

                    foreach (var layerId in layersToUnlock)
                    {
                        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                        layer.IsLocked = false;
                    }

                    foreach (var id in idsToErase)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                        if (ent == null) continue;
                        ent.Erase();
                        erasedCount++;
                    }

                    foreach (var layerId in layersToUnlock)
                    {
                        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                        layer.IsLocked = true;
                    }

                    tr.Commit();
                }

                if (erasedCount > 0)
                {
                    ed.WriteMessage($"\nErased {erasedCount} object(s).");
                    ed.Regen();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to erase entities: {ex.Message}");
            }

            return erasedCount;
        }

        internal static void TrimModelSpaceToViewportBounds()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                var viewportExtents = new List<Extents3d>();
                HashSet<ObjectId> keepIds;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var modelId = bt[BlockTableRecord.ModelSpace];
                    var model = (BlockTableRecord)tr.GetObject(modelId, OpenMode.ForRead);

                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (!btr.IsLayout || btr.Name == BlockTableRecord.ModelSpace) continue;

                        foreach (ObjectId entId in btr)
                        {
                            var vp = tr.GetObject(entId, OpenMode.ForRead) as Viewport;
                            if (vp == null) continue;
                            if (vp.Number == 1 || !vp.On) continue;

                            double viewHeight = vp.ViewHeight;
                            if (viewHeight <= 1e-9) continue;
                            double aspect = vp.Height <= 1e-9 ? 1.0 : vp.Width / vp.Height;
                            double viewWidth = viewHeight * aspect;

                            double halfW = viewWidth / 2.0;
                            double halfH = viewHeight / 2.0;
                            double twist = vp.TwistAngle;
                            double cos = Math.Cos(twist);
                            double sin = Math.Sin(twist);
                            var vx = new Vector3d(cos, sin, 0);
                            var vy = new Vector3d(-sin, cos, 0);
                            var center = vp.ViewTarget;

                            var corners = new[]
                            {
                                center + vx * halfW + vy * halfH,
                                center + vx * halfW - vy * halfH,
                                center - vx * halfW - vy * halfH,
                                center - vx * halfW + vy * halfH
                            };

                            double minX = corners.Min(c => c.X);
                            double maxX = corners.Max(c => c.X);
                            double minY = corners.Min(c => c.Y);
                            double maxY = corners.Max(c => c.Y);

                            viewportExtents.Add(new Extents3d(
                                new Point3d(minX, minY, double.NegativeInfinity),
                                new Point3d(maxX, maxY, double.PositiveInfinity)));
                        }
                    }

                    if (viewportExtents.Count == 0)
                    {
                        tr.Commit();
                        ed.WriteMessage("\nNo enabled paper space viewports were found; skipping viewport-based cleanup.");
                        return;
                    }

                    keepIds = new HashSet<ObjectId>();
                    foreach (ObjectId entId in model)
                    {
                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        var extOpt = TryGetExtents(ent);
                        if (extOpt == null)
                        {
                            continue;
                        }

                        if (viewportExtents.Any(v => ExtentsIntersectXY(extOpt.Value, v)))
                        {
                            keepIds.Add(entId);
                        }
                    }

                    tr.Commit();
                }

                if (viewportExtents.Count == 0)
                {
                    return;
                }

                if (keepIds is null)
                {
                    return;
                }

                var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                EraseEntitiesExcept(db, ed, modelSpaceId, keepIds);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nViewport clipping cleanup failed: {ex.Message}");
            }
        }

        private static bool ExtentsIntersectXY(Extents3d a, Extents3d b)
        {
            return a.MinPoint.X <= b.MaxPoint.X && a.MaxPoint.X >= b.MinPoint.X &&
                   a.MinPoint.Y <= b.MaxPoint.Y && a.MaxPoint.Y >= b.MinPoint.Y;
        }

    }
}



