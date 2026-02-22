using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        [CommandMethod("ISOWALLS", CommandFlags.Modal)]
        [CommandMethod("IW", CommandFlags.Modal)]
        public static void OnlyWalls()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            SwitchToModelSpaceViewSafe(db, ed);

            using (doc.LockDocument())
            {
                try
                {
                    int erased = 0;
                    int layersOff = 0;
                    int wallLayersFound = 0;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                        var modelSpace = (BlockTableRecord)tr.GetObject(modelSpaceId, OpenMode.ForRead);
                        var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                        var wallLayerIds = new HashSet<ObjectId>();
                        var nonWallLayerIds = new HashSet<ObjectId>();

                        foreach (ObjectId layerId in lt)
                        {
                            if (!layerId.IsValid || layerId.IsErased) continue;
                            var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                            if (ltr == null) continue;

                            if (ltr.Name.IndexOf("wall", StringComparison.OrdinalIgnoreCase) >= 0)
                                wallLayerIds.Add(layerId);
                            else
                                nonWallLayerIds.Add(layerId);
                        }

                        wallLayersFound = wallLayerIds.Count;

                        const string emptyLayerName = "0-empty-layer";
                        ObjectId emptyLayerId = ObjectId.Null;

                        if (!lt.Has(emptyLayerName))
                        {
                            lt.UpgradeOpen();
                            var newLayer = new LayerTableRecord { Name = emptyLayerName };
                            emptyLayerId = lt.Add(newLayer);
                            tr.AddNewlyCreatedDBObject(newLayer, true);
                        }
                        else
                        {
                            emptyLayerId = lt[emptyLayerName];
                        }

                        if (!emptyLayerId.IsNull)
                        {
                            var emptyLayer = (LayerTableRecord)tr.GetObject(emptyLayerId, OpenMode.ForWrite);
                            if (emptyLayer.IsOff) emptyLayer.IsOff = false;
                            if (emptyLayer.IsFrozen) emptyLayer.IsFrozen = false;
                            if (emptyLayer.IsLocked) emptyLayer.IsLocked = false;
                            db.Clayer = emptyLayerId;
                        }

                        if (wallLayerIds.Count > 0)
                        {
                            foreach (var wallLayerId in wallLayerIds)
                            {
                                var wallLayer = (LayerTableRecord)tr.GetObject(wallLayerId, OpenMode.ForWrite);
                                if (wallLayer.IsOff) wallLayer.IsOff = false;
                                if (wallLayer.IsFrozen) wallLayer.IsFrozen = false;
                            }
                        }

                        var toErase = new List<Entity>();
                        var layersToUnlock = new HashSet<ObjectId>();

                        foreach (ObjectId id in modelSpace)
                        {
                            if (!id.IsValid) continue;
                            var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                            if (ent == null) continue;

                            if (!wallLayerIds.Contains(ent.LayerId))
                            {
                                toErase.Add(ent);

                                var layer = (LayerTableRecord)tr.GetObject(ent.LayerId, OpenMode.ForRead);
                                if (layer.IsLocked)
                                    layersToUnlock.Add(ent.LayerId);
                            }
                        }

                        foreach (var layerId in layersToUnlock)
                        {
                            var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                            layer.IsLocked = false;
                        }

                        foreach (var ent in toErase)
                        {
                            if (!ent.IsErased)
                            {
                                ent.UpgradeOpen();
                                ent.Erase();
                                erased++;
                            }
                        }

                        foreach (var layerId in layersToUnlock)
                        {
                            var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                            layer.IsLocked = true;
                        }

                        foreach (var layerId in nonWallLayerIds)
                        {
                            if (layerId == db.Clayer) continue;
                            var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                            if (!ltr.IsOff)
                            {
                                ltr.IsOff = true;
                                layersOff++;
                            }
                        }

                        tr.Commit();
                    }

                    ed.WriteMessage($"\nWALLSONLY: Erased {erased} non-wall object(s). Turned off {layersOff} non-wall layer(s).");
                    if (wallLayersFound == 0)
                        ed.WriteMessage("\nWALLSONLY: No layers containing 'wall' were found.");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWALLSONLY failed: {ex.Message}");
                }
            }
        }
    }
}
