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
        // STAGE 1: Get snapshots, bind, and queue the main cleanup command.
        [CommandMethod("FINALIZE")]
        public static void FinalizeDrawingCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            bool strictProtection = StrictTitleBlockProtectionActive;
            ObjectId protectedXrefId = ProtectedTitleBlockXrefId;

            SimplerCommands.EnsureAllLayersVisibleAndUnlocked(db, ed);
            SimplerCommands.PrepareXrefLayersForCleanup(db, ed);

            _blockIdsBeforeBind.Clear();
            _originalXrefIds.Clear();
            _imageDefsToPurge.Clear();

            SimplerCommands.DetachSpecialXrefs();

            ed.WriteMessage("\n--- Stage 1: Analyzing and Binding... ---");

            try
            {
                if (strictProtection && protectedXrefId.IsNull)
                {
                    ed.WriteMessage("\nCLEANCAD strict mode failed: protected titleblock XREF was not initialized.");
                    ResetStrictTitleBlockProtection();
                    return;
                }

                int bindCount = 0;
                bool protectedFound = false;
                bool protectedResolved = false;
                bool protectedIsDwg = false;
                string protectedNameAtScan = ProtectedTitleBlockName;
                string protectedPathAtScan = ProtectedTitleBlockPath;

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    foreach (ObjectId btrId in bt)
                    {
                        _blockIdsBeforeBind.Add(btrId);
                        var btr = trans.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr == null)
                        {
                            continue;
                        }

                        if (strictProtection && btrId == protectedXrefId)
                        {
                            protectedFound = true;
                            protectedResolved = btr.IsResolved;
                            protectedIsDwg = IsDwgXrefDefinition(btr);
                            protectedNameAtScan = btr.Name ?? protectedNameAtScan;
                            protectedPathAtScan = btr.PathName ?? protectedPathAtScan;
                        }

                        if (!btr.IsFromExternalReference || !btr.IsResolved)
                        {
                            continue;
                        }

                        if (IsDwgXrefDefinition(btr))
                        {
                            _originalXrefIds.Add(btrId);
                        }
                    }

                    trans.Commit();
                }

                if (strictProtection)
                {
                    if (!protectedFound)
                    {
                        ed.WriteMessage(
                            "\nCLEANCAD strict mode failed: the protected titleblock XREF no longer exists in this drawing.");
                        ResetStrictTitleBlockProtection();
                        return;
                    }

                    if (!protectedIsDwg)
                    {
                        ed.WriteMessage(
                            $"\nCLEANCAD strict mode failed: protected titleblock '{protectedNameAtScan}' is not a DWG XREF.");
                        ResetStrictTitleBlockProtection();
                        return;
                    }

                    if (!protectedResolved)
                    {
                        ed.WriteMessage(
                            $"\nProtected titleblock '{protectedNameAtScan}' is unresolved. Attempting one XREF reload...");
                        TryReloadXrefDefinition(db, ed, protectedXrefId);
                    }

                    if (!TryGetXrefState(db, protectedXrefId, out XrefState stateAfterReload) ||
                        !stateAfterReload.Exists ||
                        !stateAfterReload.IsExternal ||
                        !stateAfterReload.IsResolved ||
                        !stateAfterReload.IsDwg)
                    {
                        string statusDetail = stateAfterReload.Exists
                            ? stateAfterReload.StatusText
                            : "Missing";
                        string pathDetail = stateAfterReload.Exists
                            ? stateAfterReload.PathName
                            : protectedPathAtScan;
                        ed.WriteMessage(
                            $"\nCLEANCAD strict mode failed: titleblock '{protectedNameAtScan}' is not bindable. Status: {statusDetail}. Path: {pathDetail}");
                        ResetStrictTitleBlockProtection();
                        return;
                    }

                    _originalXrefIds.Add(protectedXrefId);
                    ed.WriteMessage(
                        $"\nCLEANCAD strict mode: titleblock '{stateAfterReload.Name}' is resolved and queued for bind.");
                }

                ed.WriteMessage($"\nFound {_blockIdsBeforeBind.Count} total blocks and {_originalXrefIds.Count} DWG XREFs to bind.");

                if (_originalXrefIds.Count > 0)
                {
                    var idsToBind = new ObjectIdCollection(_originalXrefIds.ToArray());
                    if (!SkipBindDuringFinalize)
                    {
                        ed.WriteMessage($"\nBinding {idsToBind.Count} DWG reference(s)...");
                        db.BindXrefs(idsToBind, true);
                        bindCount = idsToBind.Count;
                    }
                    else
                    {
                        ed.WriteMessage("\nSkipping XREF binding per configuration; will detach originals only.");
                    }
                }

                if (strictProtection && !SkipBindDuringFinalize)
                {
                    if (TryGetXrefState(db, protectedXrefId, out XrefState postBindState) &&
                        postBindState.Exists &&
                        postBindState.IsExternal)
                    {
                        ed.WriteMessage(
                            $"\nCLEANCAD strict mode failed: titleblock XREF '{postBindState.Name}' remained external after bind (status: {postBindState.StatusText}).");
                        ResetStrictTitleBlockProtection();
                        return;
                    }

                    ed.WriteMessage("\nCLEANCAD strict mode: titleblock bind verification passed.");
                }

                if (bindCount > 0 || ForceDetachOriginalXrefs || _originalXrefIds.Count > 0)
                {
                    ed.WriteMessage("\nBind complete or skipped. Queueing cleanup process...");
                    doc.SendStringToExecute("_-FINALIZE-CLEANUP ", true, false, false);
                }
                else
                {
                    ed.WriteMessage("\nNo bindable DWG references found.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during bind: {ex.Message}");
                if (strictProtection)
                {
                    ResetStrictTitleBlockProtection();
                }
            }
            finally
            {
                SkipBindDuringFinalize = false;
            }
        }

        private static bool IsDwgXrefDefinition(BlockTableRecord btr)
        {
            if (btr == null) return false;

            string pathName = btr.PathName ?? string.Empty;
            if (!string.IsNullOrEmpty(pathName))
            {
                try
                {
                    if (string.Equals(Path.GetExtension(pathName), ".dwg", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                catch { }
            }

            string name = btr.Name ?? string.Empty;
            return name.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase);
        }

        private static void TryReloadXrefDefinition(Database db, Editor ed, ObjectId xrefId)
        {
            if (db == null || xrefId.IsNull) return;

            try
            {
                db.ReloadXrefs(new ObjectIdCollection(new[] { xrefId }));
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nUnable to reload protected titleblock XREF: {ex.Message}");
            }
        }

        private static bool TryGetXrefState(Database db, ObjectId xrefId, out XrefState state)
        {
            state = XrefState.Missing;
            if (db == null || xrefId.IsNull) return false;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    if (!xrefId.IsValid || xrefId.IsErased)
                    {
                        tr.Commit();
                        state = XrefState.Missing;
                        return false;
                    }

                    var btr = tr.GetObject(xrefId, OpenMode.ForRead, false, true) as BlockTableRecord;
                    if (btr == null || btr.IsErased)
                    {
                        tr.Commit();
                        state = XrefState.Missing;
                        return false;
                    }

                    state = new XrefState
                    {
                        Exists = true,
                        IsExternal = btr.IsFromExternalReference || btr.IsFromOverlayReference,
                        IsResolved = btr.IsResolved,
                        IsDwg = IsDwgXrefDefinition(btr),
                        Name = btr.Name ?? string.Empty,
                        PathName = btr.PathName ?? string.Empty,
                        StatusText = btr.XrefStatus.ToString()
                    };

                    tr.Commit();
                    return true;
                }
            }
            catch
            {
                state = XrefState.Missing;
                return false;
            }
        }

        private struct XrefState
        {
            internal bool Exists;
            internal bool IsExternal;
            internal bool IsResolved;
            internal bool IsDwg;
            internal string Name;
            internal string PathName;
            internal string StatusText;

            internal static XrefState Missing => new XrefState
            {
                Exists = false,
                IsExternal = false,
                IsResolved = false,
                IsDwg = false,
                Name = string.Empty,
                PathName = string.Empty,
                StatusText = "Missing"
            };
        }
    }
}
