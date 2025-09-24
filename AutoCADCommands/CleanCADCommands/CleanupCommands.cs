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

        /// <summary>
        /// Scans all spaces for block references matching title block names and logs their status.
        /// </summary>
        public static void LogTitleBlockReferenceStatus(Database db, Editor ed, string stage)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    string[] titleBlockHints = { "x-tb", "title", "tblock", "border", "sheet" };
                    var tbDefIds = new HashSet<ObjectId>();

                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // Find all block definitions that look like a title block
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr == null || btr.IsErased) continue;

                        string nameLower = (btr.Name ?? string.Empty).ToLowerInvariant();
                        string pathLower = string.Empty;
                        if (btr.IsFromExternalReference)
                        {
                            try { pathLower = System.IO.Path.GetFileNameWithoutExtension(btr.PathName ?? string.Empty).ToLowerInvariant(); } catch { }
                        }

                        if (titleBlockHints.Any(h => nameLower.Contains(h) || (!string.IsNullOrEmpty(pathLower) && pathLower.Contains(h))))
                        {
                            tbDefIds.Add(btrId);
                        }
                    }

                    if (tbDefIds.Count == 0)
                    {
                        ed.WriteMessage($"\n[LOG @ {stage}] No title block DEFINITION found.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n[LOG @ {stage}] Found {tbDefIds.Count} possible title block definition(s). Checking for references...");

                    int refCount = 0;
                    // Now search all spaces for references to these definitions
                    foreach (ObjectId btrId in bt)
                    {
                        var space = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        // *** FIX IS HERE: Correctly check if the BTR is a layout (Model or Paper) ***
                        if (space == null || !space.IsLayout) continue;

                        foreach (ObjectId entId in space)
                        {
                            if (entId.IsErased) continue;
                            var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;

                            ObjectId effectiveDefId = br.BlockTableRecord;
                            if (br.IsDynamicBlock)
                            {
                                try { effectiveDefId = br.DynamicBlockTableRecord; } catch { }
                            }

                            if (tbDefIds.Contains(effectiveDefId))
                            {
                                refCount++;
                                var def = (BlockTableRecord)tr.GetObject(effectiveDefId, OpenMode.ForRead);
                                ed.WriteMessage($"\n[LOG @ {stage}]   -> FOUND Title Block Reference. Handle: {br.Handle}, Definition: '{def.Name}', IsXref: {def.IsFromExternalReference}, Space: '{space.Name}'.");
                            }
                        }
                    }

                    if (refCount == 0)
                    {
                        ed.WriteMessage($"\n[LOG @ {stage}]   -> !!! NO Title Block Reference FOUND in any space. !!!");
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[LOG @ {stage}] Error during logging: {ex.Message}");
            }
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