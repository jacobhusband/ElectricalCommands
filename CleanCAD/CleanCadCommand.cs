using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        // Orchestrator: run DETACHSPECIALXREFS, then EMBEDFROMXREFS, then FINALIZE
        [CommandMethod("CLEANCAD", CommandFlags.Modal)]
        public static void RunCleanCad()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            try
            {
                ed.WriteMessage("\n--- CleanCAD: Detaching special refs, embedding, and finalizing ---");
                // Step 1: detach targeted image/DWG xrefs, delete matching blocks, freeze layers
                DetachSpecialXrefs();

                // Step 2: embed images over raster xrefs; when done, automatically run FINALIZE
                _chainFinalizeAfterEmbed = true;
                EmbedFromXrefs();
                // Note: FINALIZE is triggered automatically when embedding completes or immediately if none found.
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nCleanCAD failed: {ex.Message}");
                _chainFinalizeAfterEmbed = false;
            }
        }
    }
}

