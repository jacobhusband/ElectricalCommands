using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
    public partial class GeneralCommands
    {
        private enum TurnRedMode
        {
            Cancel = 0,
            Xref = 1,
            Layer = 2,
            Region = 3
        }

        [CommandMethod("TURNRED", CommandFlags.UsePickSet)]
        public static void TurnLayerRed()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            TurnRedMode mode = PromptTurnRedMode(ed);
            if (mode == TurnRedMode.Cancel) return;

            switch (mode)
            {
                case TurnRedMode.Xref:
                    TurnRedXref(db, ed);
                    break;
                case TurnRedMode.Region:
                    TurnRedRegion(db, ed);
                    break;
                default:
                    TurnRedLayer(db, ed);
                    break;
            }
        }

        private static string TryGetEntityLayerName(ObjectId objectId, Database hostDb, Editor ed)
        {
            try
            {
                Database entityDb = objectId.Database ?? hostDb;
                using (Transaction tr = entityDb.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                    string layer = ent?.Layer;
                    tr.Commit();
                    return layer;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to read selected object: {ex.Message}");
                return null;
            }
        }

        private static TurnRedMode PromptTurnRedMode(Editor ed)
        {
            PromptKeywordOptions pko =
                new PromptKeywordOptions(
                    "\nTURNRED option [Xref/Layer/Region] <Layer>: ",
                    "Xref Layer Region");
            pko.AllowNone = true;

            PromptResult pr = ed.GetKeywords(pko);
            if (pr.Status == PromptStatus.None)
                return TurnRedMode.Layer;
            if (pr.Status != PromptStatus.OK)
                return TurnRedMode.Cancel;

            string keyword = pr.StringResult ?? string.Empty;
            if (keyword.Equals("Xref", StringComparison.OrdinalIgnoreCase))
                return TurnRedMode.Xref;
            if (keyword.Equals("Region", StringComparison.OrdinalIgnoreCase))
                return TurnRedMode.Region;

            return TurnRedMode.Layer;
        }

        private static void TurnRedLayer(Database db, Editor ed)
        {
            ObjectId selectedId = ObjectId.Null;
            ObjectId[] containers = Array.Empty<ObjectId>();

            PromptSelectionResult psr = ed.SelectImplied();
            if (psr.Status == PromptStatus.OK && psr.Value.Count > 0)
            {
                selectedId = psr.Value[0].ObjectId;
            }
            else
            {
                PromptNestedEntityOptions pneo =
                    new PromptNestedEntityOptions(
                        "\nSelect an object to change its layer color to red: ");
                PromptNestedEntityResult pner = ed.GetNestedEntity(pneo);
                if (pner.Status == PromptStatus.OK)
                {
                    selectedId = pner.ObjectId;
                    containers = pner.GetContainers();
                }
                else if (pner.Status == PromptStatus.Error)
                {
                    PromptEntityOptions peo =
                        new PromptEntityOptions(
                            "\nSelect an object to change its layer color to red: ");
                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;
                    selectedId = per.ObjectId;
                }
                else
                {
                    return;
                }
            }

            if (selectedId == ObjectId.Null) return;

            string entityLayer = TryGetEntityLayerName(selectedId, db, ed);
            if (string.IsNullOrWhiteSpace(entityLayer))
            {
                ed.WriteMessage("\nUnable to determine the layer of the selected object.");
                return;
            }

            List<string> xrefNames = ExtractXrefNames(containers);
            string resolvedLayerName = ResolveLayerNameForSelection(db, ed, entityLayer, xrefNames);
            if (string.IsNullOrWhiteSpace(resolvedLayerName))
            {
                if (xrefNames.Count > 0)
                    ed.WriteMessage($"\nLayer '{string.Join("|", xrefNames)}|{entityLayer}' was not found in this drawing.");
                else
                    ed.WriteMessage($"\nLayer '{entityLayer}' was not found in this drawing.");
                return;
            }

            if (SetLayersColor(db, ed, new[] { resolvedLayerName }, 1, out int changed) && changed > 0)
            {
                ed.WriteMessage($"\nLayer '{resolvedLayerName}' set to red.");
                ed.Regen();
            }
        }

        private static void TurnRedXref(Database db, Editor ed)
        {
            ObjectId selectedId = ObjectId.Null;
            ObjectId[] containers = Array.Empty<ObjectId>();

            PromptSelectionResult psr = ed.SelectImplied();
            if (psr.Status == PromptStatus.OK && psr.Value.Count > 0)
            {
                selectedId = psr.Value[0].ObjectId;
            }
            else
            {
                PromptNestedEntityOptions pneo =
                    new PromptNestedEntityOptions("\nSelect an object in the XREF: ");
                PromptNestedEntityResult pner = ed.GetNestedEntity(pneo);
                if (pner.Status == PromptStatus.OK)
                {
                    selectedId = pner.ObjectId;
                    containers = pner.GetContainers();
                }
                else if (pner.Status == PromptStatus.Error)
                {
                    PromptEntityOptions peo =
                        new PromptEntityOptions("\nSelect an XREF block reference: ");
                    peo.SetRejectMessage("\nSelected object is not a block reference.");
                    peo.AddAllowedClass(typeof(BlockReference), true);
                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;
                    selectedId = per.ObjectId;
                }
                else
                {
                    return;
                }
            }

            if (selectedId == ObjectId.Null) return;

            if (!TryResolveXrefNameFromSelection(selectedId, containers, out string xrefName))
            {
                ed.WriteMessage("\nThe selected object is not part of an XREF.");
                return;
            }

            int layersChanged = SetXrefLayersColor(db, ed, xrefName, 1);
            if (layersChanged > 0)
            {
                ed.WriteMessage($"\nChanged {layersChanged} layer(s) in XREF '{xrefName}' to red.");
                ed.Regen();
            }
            else
            {
                ed.WriteMessage($"\nNo layers found for XREF '{xrefName}'.");
            }
        }

        private static void TurnRedRegion(Database db, Editor ed)
        {
            PromptPointResult ppr1 = ed.GetPoint("\nFirst corner of region: ");
            if (ppr1.Status != PromptStatus.OK) return;

            PromptCornerOptions pco =
                new PromptCornerOptions("\nOther corner: ", ppr1.Value);
            PromptPointResult ppr2 = ed.GetCorner(pco);
            if (ppr2.Status != PromptStatus.OK) return;

            PromptSelectionResult psr = ed.SelectCrossingWindow(ppr1.Value, ppr2.Value);
            if (psr.Status != PromptStatus.OK || psr.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects found in the selected region.");
                return;
            }

            int changed = 0;
            int xrefBlocks = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in psr.Value)
                {
                    if (so == null || so.ObjectId == ObjectId.Null) continue;

                    Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    if (ent is BlockReference br)
                    {
                        BlockTableRecord btr =
                            tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null && btr.IsFromExternalReference)
                        {
                            xrefBlocks++;
                            continue;
                        }
                    }

                    try
                    {
                        ent.UpgradeOpen();
                        ent.Color = Color.FromColorIndex(ColorMethod.ByAci, 1); // Red
                        changed++;
                    }
                    catch
                    {
                        // Skip entities that cannot be modified (e.g., locked layers).
                    }
                }

                tr.Commit();
            }

            if (changed == 0 && xrefBlocks == 0)
            {
                ed.WriteMessage("\nNo objects were changed.");
                return;
            }

            if (changed > 0)
                ed.WriteMessage($"\nChanged {changed} object(s) to red.");

            if (xrefBlocks > 0)
            {
                ed.WriteMessage(
                    "\nNote: XREF objects inside the region are not editable from the host drawing. " +
                    "Those nested entities were not recolored.");
            }
            ed.Regen();
        }

        private static List<string> ExtractXrefNames(ObjectId[] containers)
        {
            var result = new List<string>();
            if (containers == null || containers.Length == 0)
                return result;

            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId containerId in containers)
            {
                if (!TryGetXrefName(containerId, out string xrefName))
                    continue;

                if (string.IsNullOrWhiteSpace(xrefName)) continue;
                if (seen.Add(xrefName))
                    result.Add(xrefName);
            }

            return result;
        }

        private static bool TryResolveXrefNameFromSelection(
            ObjectId selectedId,
            ObjectId[] containers,
            out string xrefName)
        {
            xrefName = null;

            if (TryGetXrefName(selectedId, out string directName))
            {
                xrefName = directName;
                return true;
            }

            if (containers == null || containers.Length == 0)
                return false;

            foreach (ObjectId containerId in containers)
            {
                if (TryGetXrefName(containerId, out string containerName))
                {
                    xrefName = containerName;
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetXrefName(ObjectId containerId, out string xrefName)
        {
            xrefName = null;
            Database containerDb = containerId.Database;
            if (containerDb == null)
                return false;

            bool isXref = false;

            using (Transaction tr = containerDb.TransactionManager.StartTransaction())
            {
                BlockReference br = tr.GetObject(containerId, OpenMode.ForRead) as BlockReference;
                if (br != null)
                {
                    BlockTableRecord btr =
                        tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null && btr.IsFromExternalReference)
                    {
                        xrefName = btr.Name;
                        isXref = true;
                    }
                }

                tr.Commit();
            }

            return isXref;
        }

        private static string ResolveLayerNameForSelection(
            Database db,
            Editor ed,
            string entityLayer,
            List<string> xrefNames)
        {
            if (string.IsNullOrWhiteSpace(entityLayer))
                return null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (lt == null)
                    {
                        ed.WriteMessage("\nUnable to access the layer table.");
                        return null;
                    }

                    if (xrefNames == null || xrefNames.Count == 0)
                    {
                        if (lt.Has(entityLayer))
                            return entityLayer;
                    }
                    else
                    {
                        string primary = string.Join("|", xrefNames) + "|" + entityLayer;
                        if (lt.Has(primary))
                            return primary;

                        var reversed = new List<string>(xrefNames);
                        reversed.Reverse();
                        string secondary = string.Join("|", reversed) + "|" + entityLayer;
                        if (lt.Has(secondary))
                            return secondary;

                        string best = null;
                        int bestScore = -1;
                        int bestExtra = int.MaxValue;

                        foreach (ObjectId layerId in lt)
                        {
                            LayerTableRecord ltr = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                            if (ltr == null) continue;

                            string name = ltr.Name;
                            if (!name.EndsWith("|" + entityLayer, StringComparison.OrdinalIgnoreCase))
                                continue;

                            string[] parts = name.Split('|');
                            if (parts.Length < 2) continue;

                            var prefixParts = new List<string>(parts.Length - 1);
                            for (int i = 0; i < parts.Length - 1; i++)
                                prefixParts.Add(parts[i]);

                            int score = ScorePrefixMatch(prefixParts, xrefNames);
                            if (score <= 0) continue;

                            int extra = Math.Abs(prefixParts.Count - xrefNames.Count);
                            if (score > bestScore || (score == bestScore && extra < bestExtra))
                            {
                                bestScore = score;
                                bestExtra = extra;
                                best = name;
                            }
                        }

                        return best;
                    }

                    return null;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred: {ex.Message}");
                    return null;
                }
            }
        }

        private static int ScorePrefixMatch(List<string> prefixParts, List<string> xrefNames)
        {
            if (prefixParts == null || xrefNames == null || xrefNames.Count == 0)
                return 0;

            int forward = CountOrderedMatches(prefixParts, xrefNames);

            var reversed = new List<string>(xrefNames);
            reversed.Reverse();
            int backward = CountOrderedMatches(prefixParts, reversed);

            return Math.Max(forward, backward);
        }

        private static int CountOrderedMatches(List<string> prefixParts, List<string> xrefNames)
        {
            int score = 0;
            int idx = 0;

            foreach (string part in prefixParts)
            {
                if (idx >= xrefNames.Count) break;
                if (string.Equals(part, xrefNames[idx], StringComparison.OrdinalIgnoreCase))
                {
                    score++;
                    idx++;
                }
            }

            return score;
        }

        private static bool SetLayersColor(
            Database db,
            Editor ed,
            IEnumerable<string> layerNames,
            short colorIndex,
            out int changedCount)
        {
            changedCount = 0;
            if (layerNames == null) return false;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (lt == null)
                    {
                        ed.WriteMessage("\nUnable to access the layer table.");
                        return false;
                    }

                    foreach (string layerName in layerNames)
                    {
                        if (string.IsNullOrWhiteSpace(layerName)) continue;
                        if (!lt.Has(layerName)) continue;

                        LayerTableRecord ltr =
                            tr.GetObject(lt[layerName], OpenMode.ForWrite) as LayerTableRecord;
                        if (ltr == null) continue;

                        ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                        changedCount++;
                    }

                    tr.Commit();
                    return true;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred: {ex.Message}");
                    return false;
                }
            }
        }

        private static int SetXrefLayersColor(Database db, Editor ed, string xrefName, short colorIndex)
        {
            if (string.IsNullOrWhiteSpace(xrefName))
                return 0;

            int layersChanged = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (lt == null)
                    {
                        ed.WriteMessage("\nUnable to access the layer table.");
                        return 0;
                    }

                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        if (ltr == null) continue;

                        if (ltr.Name.StartsWith(xrefName + "|", StringComparison.InvariantCultureIgnoreCase))
                        {
                            ltr.UpgradeOpen();
                            ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                            layersChanged++;
                        }
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred: {ex.Message}");
                    return 0;
                }
            }

            return layersChanged;
        }
    }
}
