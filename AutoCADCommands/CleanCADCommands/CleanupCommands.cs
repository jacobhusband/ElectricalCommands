using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
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
        internal static bool RunRemoveRemainingAfterFinalize = false;
        internal static bool StrictTitleBlockProtectionActive = false;
        internal static bool StrictTitleBlockBindFailed = false;
        internal static bool AbortRemainingXrefDetach = false;
        internal static ObjectId ProtectedTitleBlockXrefId = ObjectId.Null;
        internal static string ProtectedTitleBlockName = string.Empty;
        internal static string ProtectedTitleBlockPath = string.Empty;
        internal static string ProtectedTitleBlockCanonicalName = string.Empty;
        internal static string ProtectedTitleBlockFileName = string.Empty;
        internal static string ProtectedTitleBlockLayoutName = string.Empty;

        internal static void EnableStrictTitleBlockProtection(ObjectId xrefId, string blockName, string pathName, string layoutName = "")
        {
            StrictTitleBlockProtectionActive = !xrefId.IsNull;
            StrictTitleBlockBindFailed = false;
            AbortRemainingXrefDetach = false;
            ProtectedTitleBlockXrefId = xrefId;
            ProtectedTitleBlockName = blockName ?? string.Empty;
            ProtectedTitleBlockPath = pathName ?? string.Empty;

            string canonicalFromBlock = CanonicalizeXrefToken(blockName);
            string canonicalFromPath = CanonicalizeXrefToken(Path.GetFileName(pathName ?? string.Empty));
            ProtectedTitleBlockCanonicalName = !string.IsNullOrWhiteSpace(canonicalFromBlock)
                ? canonicalFromBlock
                : canonicalFromPath;
            ProtectedTitleBlockFileName = canonicalFromPath;
            ProtectedTitleBlockLayoutName = (layoutName ?? string.Empty).Trim();
        }

        internal static void ResetStrictTitleBlockProtection()
        {
            StrictTitleBlockProtectionActive = false;
            StrictTitleBlockBindFailed = false;
            AbortRemainingXrefDetach = false;
            ProtectedTitleBlockXrefId = ObjectId.Null;
            ProtectedTitleBlockName = string.Empty;
            ProtectedTitleBlockPath = string.Empty;
            ProtectedTitleBlockCanonicalName = string.Empty;
            ProtectedTitleBlockFileName = string.Empty;
            ProtectedTitleBlockLayoutName = string.Empty;
        }

        internal static bool IsProtectedTitleBlockXref(ObjectId xrefId)
        {
            return StrictTitleBlockProtectionActive &&
                   !ProtectedTitleBlockXrefId.IsNull &&
                   xrefId == ProtectedTitleBlockXrefId;
        }

        internal static void MarkStrictTitleBlockBindFailure()
        {
            StrictTitleBlockBindFailed = true;
            AbortRemainingXrefDetach = true;
        }

        internal static string CanonicalizeXrefToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string token = value.Trim().ToLowerInvariant();
            token = token.Replace("\\", "/");
            try
            {
                token = Path.GetFileNameWithoutExtension(token) ?? token;
            }
            catch { }
            return token.Trim();
        }

        internal static string NormalizeXrefPathToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return value.Trim().Trim('"').Replace('/', '\\');
        }

        private static string GetDrawingDirectoryForXrefComparison(Database db)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null && !string.IsNullOrWhiteSpace(doc.Name))
                {
                    string docDir = Path.GetDirectoryName(doc.Name);
                    if (!string.IsNullOrWhiteSpace(docDir)) return docDir;
                }
            }
            catch { }

            try
            {
                if (db != null && !string.IsNullOrWhiteSpace(db.Filename))
                {
                    string dbDir = Path.GetDirectoryName(db.Filename);
                    if (!string.IsNullOrWhiteSpace(dbDir)) return dbDir;
                }
            }
            catch { }

            return string.Empty;
        }

        internal static string ResolveXrefPathForComparison(Database db, string rawPath)
        {
            string normalized = NormalizeXrefPathToken(rawPath);
            if (string.IsNullOrWhiteSpace(normalized)) return string.Empty;

            try
            {
                if (Path.IsPathRooted(normalized))
                {
                    return Path.GetFullPath(normalized);
                }
            }
            catch { }

            try
            {
                string drawingDir = GetDrawingDirectoryForXrefComparison(db);
                if (!string.IsNullOrWhiteSpace(drawingDir))
                {
                    return Path.GetFullPath(Path.Combine(drawingDir, normalized));
                }
            }
            catch { }

            try
            {
                if (db != null)
                {
                    string nameOnly = Path.GetFileName(normalized);
                    if (!string.IsNullOrWhiteSpace(nameOnly))
                    {
                        string found = HostApplicationServices.Current?.FindFile(nameOnly, db, FindFileHint.Default);
                        if (!string.IsNullOrWhiteSpace(found))
                        {
                            return Path.GetFullPath(found);
                        }
                    }
                }
            }
            catch { }

            return normalized;
        }

        internal static bool AreEquivalentXrefPaths(Database db, string leftPath, string rightPath)
        {
            string leftRaw = NormalizeXrefPathToken(leftPath);
            string rightRaw = NormalizeXrefPathToken(rightPath);
            if (string.IsNullOrWhiteSpace(leftRaw) || string.IsNullOrWhiteSpace(rightRaw))
            {
                return false;
            }

            if (string.Equals(leftRaw, rightRaw, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string leftResolved = ResolveXrefPathForComparison(db, leftRaw);
            string rightResolved = ResolveXrefPathForComparison(db, rightRaw);
            if (!string.IsNullOrWhiteSpace(leftResolved) &&
                !string.IsNullOrWhiteSpace(rightResolved) &&
                string.Equals(leftResolved, rightResolved, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string leftFile = CanonicalizeXrefToken(Path.GetFileName(leftRaw));
            string rightFile = CanonicalizeXrefToken(Path.GetFileName(rightRaw));
            return !string.IsNullOrWhiteSpace(leftFile) &&
                   string.Equals(leftFile, rightFile, StringComparison.OrdinalIgnoreCase);
        }

        internal static bool IsProtectedTitleBlockFingerprintMatch(Database db, string blockName, string pathName)
        {
            if (!StrictTitleBlockProtectionActive) return false;

            string blockCanonical = CanonicalizeXrefToken(blockName);
            string pathCanonical = CanonicalizeXrefToken(Path.GetFileName(pathName ?? string.Empty));

            bool canonicalMatch = !string.IsNullOrWhiteSpace(ProtectedTitleBlockCanonicalName) &&
                (string.Equals(blockCanonical, ProtectedTitleBlockCanonicalName, StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(pathCanonical, ProtectedTitleBlockCanonicalName, StringComparison.OrdinalIgnoreCase));

            bool fileMatch = !string.IsNullOrWhiteSpace(ProtectedTitleBlockFileName) &&
                (string.Equals(pathCanonical, ProtectedTitleBlockFileName, StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(blockCanonical, ProtectedTitleBlockFileName, StringComparison.OrdinalIgnoreCase));

            bool pathMatch = AreEquivalentXrefPaths(db, pathName, ProtectedTitleBlockPath);

            return canonicalMatch || fileMatch || pathMatch;
        }

        internal static bool IsProtectedTitleBlockFingerprintMatch(string blockName, string pathName)
        {
            return IsProtectedTitleBlockFingerprintMatch(null, blockName, pathName);
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
    }
}
