using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
    public partial class GeneralCommands
    {
        [CommandMethod("LAYGREEN", CommandFlags.UsePickSet)]
        [CommandMethod("LG", CommandFlags.UsePickSet)]
        public static void TurnXrefLayersGreen()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ObjectId xrefId = ObjectId.Null;

            // First, check for a pre-selected (pickfirst) entity
            PromptSelectionResult psr = ed.SelectImplied();

            if (psr.Status == PromptStatus.OK && psr.Value.Count > 0)
            {
                // If something was pre-selected, take the first object ID
                xrefId = psr.Value[0].ObjectId;
            }
            else
            {
                // If nothing was pre-selected, prompt the user to select an XREF
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect an XREF:");
                peo.SetRejectMessage("\nSelected object is not a block reference.");
                peo.AddAllowedClass(typeof(BlockReference), true);
                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK)
                    return;

                xrefId = per.ObjectId;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the selected block reference
                    BlockReference blockRef = tr.GetObject(xrefId, OpenMode.ForRead) as BlockReference;

                    if (blockRef != null)
                    {
                        // Get the block table record of the selected block reference
                        BlockTableRecord btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                        // Check if it is an XREF
                        if (btr != null && btr.IsFromExternalReference)
                        {
                            string xrefName = btr.Name;
                            ed.WriteMessage($"\nProcessing XREF: {xrefName}");

                            // AutoCAD can expose a nested XREF's dependent layers under
                            // the nested reference name rather than only under the selected
                            // top-level prefix. Walk the XREF graph so every descendant is
                            // included, while the visited-name set protects against cycles.
                            HashSet<string> xrefNames = GetXrefHierarchyNames(
                                db,
                                btr.ObjectId,
                                xrefName);

                            // Get the LayerTable
                            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                            int layersChanged = 0;

                            // Iterate through all the layers in the drawing
                            foreach (ObjectId layerId in lt)
                            {
                                LayerTableRecord ltr = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;

                                // XREF layer names are prefixed with "XREF_NAME|". Checking
                                // every name in the selected hierarchy also covers nested
                                // references whose layers are not qualified by the root name.
                                if (ltr != null && IsLayerInXrefHierarchy(ltr.Name, xrefNames))
                                {
                                    // Upgrade the layer to write access
                                    ltr.UpgradeOpen();

                                    // Set the color to green
                                    ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 3); // 3 is the ACI for green
                                    layersChanged++;
                                }
                            }

                            if (layersChanged > 0)
                            {
                                int nestedXrefCount = Math.Max(0, xrefNames.Count - 1);
                                ed.WriteMessage(
                                    $"\nChanged the color of {layersChanged} layer(s) in the selected XREF" +
                                    (nestedXrefCount > 0
                                        ? $" and {nestedXrefCount} nested XREF(s) to green."
                                        : " to green."));
                            }
                            else
                            {
                                ed.WriteMessage("\nNo layers found for the selected XREF.");
                            }
                        }
                        else
                        {
                            ed.WriteMessage("\nThe selected block is not an XREF.");
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nThe selected object is not a valid block reference.");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred: {ex.Message}");
                }

                // Commit the changes to the database
                tr.Commit();
            }

            // Regenerate the drawing to show the color changes immediately
            ed.Regen();
        }

        private static HashSet<string> GetXrefHierarchyNames(
            Database db,
            ObjectId rootBtrId,
            string rootXrefName)
        {
            var xrefNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (XrefGraph xrefGraph = db.GetHostDwgXrefGraph(false))
                {
                    XrefGraphNode rootNode = xrefGraph.GetXrefNode(rootBtrId);
                    CollectXrefHierarchyNames(rootNode, xrefNames);
                }
            }
            catch (System.Exception)
            {
                // Fall back to the original top-level behavior if AutoCAD
                // cannot construct an XREF graph for this drawing.
            }

            if (!string.IsNullOrWhiteSpace(rootXrefName))
                xrefNames.Add(rootXrefName);

            return xrefNames;
        }

        private static void CollectXrefHierarchyNames(
            XrefGraphNode node,
            HashSet<string> xrefNames)
        {
            if (node == null || string.IsNullOrWhiteSpace(node.Name) || !xrefNames.Add(node.Name))
                return;

            for (int index = 0; index < node.NumOut; index++)
            {
                CollectXrefHierarchyNames(node.Out(index) as XrefGraphNode, xrefNames);
            }
        }

        private static bool IsLayerInXrefHierarchy(
            string layerName,
            IEnumerable<string> xrefNames)
        {
            if (string.IsNullOrWhiteSpace(layerName))
                return false;

            foreach (string xrefName in xrefNames)
            {
                if (layerName.StartsWith(xrefName + "|", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }
    }
}
