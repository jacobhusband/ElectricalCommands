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
    }
}

