using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;

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

                            // Get the LayerTable
                            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                            int layersChanged = 0;

                            // Iterate through all the layers in the drawing
                            foreach (ObjectId layerId in lt)
                            {
                                LayerTableRecord ltr = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;

                                // XREF layer names are typically prefixed with "XREF_NAME|"
                                if (ltr != null && ltr.Name.StartsWith(xrefName + "|", System.StringComparison.InvariantCultureIgnoreCase))
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
                                ed.WriteMessage($"\nChanged the color of {layersChanged} layer(s) in the selected XREF to green.");
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
    }
}
