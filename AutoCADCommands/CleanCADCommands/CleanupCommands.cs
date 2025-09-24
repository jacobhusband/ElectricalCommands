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

        [CommandMethod("ZOOMTOTBPS", CommandFlags.Modal)]
        public static void ZoomToTitleBlockInPaperSpace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // Ensure we are in paper space
                if (db.TileMode)
                {
                    Application.SetSystemVariable("TILEMODE", 0);
                }

                Extents3d? tbExtents = null;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lm = LayoutManager.Current;
                    string currentLayout = lm.CurrentLayout;
                    var layoutId = lm.GetLayoutId(currentLayout);
                    var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                    tbExtents = GetTitleBlockUnionExtentsInPaper_Robust(btr, tr);
                    tr.Commit();
                }

                if (tbExtents.HasValue)
                {
                    ed.WriteMessage("\nZooming to paper space title block...");
                    ZoomToTitleBlock(ed, new[] { tbExtents.Value.MinPoint, tbExtents.Value.MaxPoint });
                }
                else
                {
                    ed.WriteMessage("\nCould not find a title block to zoom to in the current layout.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during zoom to paper space title block: {ex.Message}");
            }
        }

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
        private static bool ExtentsIntersectXY(Extents3d a, Extents3d b)
        {
            return a.MinPoint.X <= b.MaxPoint.X && a.MaxPoint.X >= b.MinPoint.X &&
                   a.MinPoint.Y <= b.MaxPoint.Y && a.MaxPoint.Y >= b.MinPoint.Y;
        }

        private static bool ExtentsContainsXY(Extents3d container, Extents3d item)
        {
            return container.MinPoint.X <= item.MinPoint.X &&
                   container.MinPoint.Y <= item.MinPoint.Y &&
                   container.MaxPoint.X >= item.MaxPoint.X &&
                   container.MaxPoint.Y >= item.MaxPoint.Y;
        }
    }
}