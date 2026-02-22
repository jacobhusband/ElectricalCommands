using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace ElectricalCommands
{
    /// <summary>
    /// Standalone plot commands bundle:
    /// - P22
    /// - P24
    /// - P30
    /// - PLOTMOVE
    /// </summary>
    public class PlotCommands
    {
        /// <summary>
        /// Core helper used by P22/P24/P30 to drive -PLOT for a specific paper size.
        /// This mirrors the original CleanCADCommands implementation semantics.
        /// </summary>
        /// <param name="paperSize">Paper size string exactly as recognized by AutoCAD's "DWG to PDF.pc3".</param>
        /// <param name="message">User-facing message written before plotting.</param>
        private static void PlotToPdf(string paperSize, string message)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            ed.WriteMessage(message);

            // Basic scripted -PLOT sequence. This assumes:
            // - current layout is active
            // - user/window selection context where applicable
            doc.SendStringToExecute("-PLOT\n", false, false, true);
            doc.SendStringToExecute("Y\n", false, false, true);                 // Detailed plot configuration
            doc.SendStringToExecute("\n", false, false, true);                  // Current layout
            doc.SendStringToExecute("DWG to PDF.pc3\n", false, false, true);    // Device
            doc.SendStringToExecute(paperSize + "\n", false, false, true);      // Paper size
            doc.SendStringToExecute("I\n", false, false, true);                 // Inches
            doc.SendStringToExecute("L\n", false, false, true);                 // Landscape
            doc.SendStringToExecute("N\n", false, false, true);                 // Plot upside down? No
            doc.SendStringToExecute("L\n", false, false, true);                 // Window (user/window selection will follow)
            // Window corners are interactive / context-driven; we do not hard-code here
            doc.SendStringToExecute("1:1\n", false, false, true);               // Scale
            doc.SendStringToExecute("0.00,0.00\n", false, false, true);         // Plot offset
            doc.SendStringToExecute("Y\n", false, false, true);                 // Plot with plot styles
            doc.SendStringToExecute("510-monochrome.ctb\n", false, false, true);// Plot style table
            doc.SendStringToExecute("Y\n", false, false, true);                 // Plot with lineweights
            doc.SendStringToExecute("N\n", false, false, true);                 // Scale lineweights with plot scale? No
            doc.SendStringToExecute("N\n", false, false, true);                 // Plot paper space last? No
            doc.SendStringToExecute("N\n", false, false, true);                 // Hide paper space objects? No
            doc.SendStringToExecute("\n", false, false, true);                  // Save changes to page setup
            doc.SendStringToExecute("Y\n", false, false, true);                 // Write the plot to a file
        }

        [CommandMethod("P30", CommandFlags.Modal)]
        public static void Plot30x42()
        {
            PlotToPdf("ARCH full bleed E1 (30.00 x 42.00 Inches)", "\nStarting 30x42 PDF plot sequence...");
        }

        [CommandMethod("P24", CommandFlags.Modal)]
        public static void Plot24x36()
        {
            PlotToPdf("ARCH full bleed D (24.00 x 36.00 Inches)", "\nStarting 24x36 PDF plot sequence...");
        }

        [CommandMethod("P22", CommandFlags.Modal)]
        public static void Plot22x34()
        {
            PlotToPdf("ANSI full bleed D (22.00 x 34.00 Inches)", "\nStarting 22x34 PDF plot sequence...");
        }

        /// <summary>
        /// PLOTMOVE:
        /// 1. Prompts for a point.
        /// 2. Runs a scripted 30x42 plot to PDF (DWG to PDF.pc3).
        /// 3. Moves all entities in the current space so that the selected point ends up at (0,0,0).
        /// </summary>
        [CommandMethod("PLOTMOVE", CommandFlags.Modal)]
        [CommandMethod("QM", CommandFlags.Modal)]
        public static void PlotAndMove()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // 1) Prompt before plotting so we know the reference point.
            PromptPointResult ppr = ed.GetPoint("\nSelect a point to move to origin: ");
            if (ppr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nPoint selection cancelled.");
                return;
            }
            Point3d sourcePoint = ppr.Value;

            // 2) Fire the plot sequence (mirrors the original PLOTMOVE implementation).
            ed.WriteMessage("\nStarting plot sequence...");
            doc.SendStringToExecute("-PLOT\n", false, false, true);
            doc.SendStringToExecute("Y\n", false, false, true);                             // Detailed plot configuration
            doc.SendStringToExecute("\n", false, false, true);                              // Current layout
            doc.SendStringToExecute("DWG to PDF.pc3\n", false, false, true);                // Device
            doc.SendStringToExecute("ARCH full bleed E1 (30.00 x 42.00 Inches)\n", false, false, true);
            doc.SendStringToExecute("I\n", false, false, true);                             // Inches
            doc.SendStringToExecute("L\n", false, false, true);                             // Landscape
            doc.SendStringToExecute("N\n", false, false, true);                             // Plot upside down? No
            doc.SendStringToExecute("W\n", false, false, true);                             // Window
            doc.SendStringToExecute("0.00,0.00\n", false, false, true);                     // Lower-left corner
            doc.SendStringToExecute("1000000,1000000\n", false, false, true);               // Upper-right corner
            doc.SendStringToExecute("1:1\n", false, false, true);                           // Scale
            doc.SendStringToExecute("0.00,0.00\n", false, false, true);                     // Plot offset
            doc.SendStringToExecute("Y\n", false, false, true);                             // Plot with plot styles
            doc.SendStringToExecute("510-monochrome.ctb\n", false, false, true);            // Plot style table
            doc.SendStringToExecute("Y\n", false, false, true);                             // Plot with lineweights
            doc.SendStringToExecute("N\n", false, false, true);                             // Scale lineweights
            doc.SendStringToExecute("N\n", false, false, true);                             // Plot paper space last
            doc.SendStringToExecute("N\n", false, false, true);                             // Hide paper space objects
            doc.SendStringToExecute("\n", false, false, true);                              // Save page setup
            doc.SendStringToExecute("Y\n", false, false, true);                             // Plot to file
            doc.SendStringToExecute("\n", false, false, true);                              // Accept default filename
            doc.SendStringToExecute("\n", false, false, true);                              // Final confirmation

            // 3) Move all entities in current space so sourcePoint maps to origin.
            ed.WriteMessage("\nMoving objects...");
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                    Point3d origin = new Point3d(0, 0, 0);
                    Vector3d vector = sourcePoint.GetVectorTo(origin);
                    Matrix3d displacement = Matrix3d.Displacement(vector);

                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                        if (ent == null) continue;
                        ent.TransformBy(displacement);
                    }

                    tr.Commit();
                    ed.WriteMessage("\nAll content moved successfully.");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError moving content: {ex.Message}");
                    tr.Abort();
                }
            }
        }
    }
}
