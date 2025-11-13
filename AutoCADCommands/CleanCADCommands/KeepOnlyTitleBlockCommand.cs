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
    }
}