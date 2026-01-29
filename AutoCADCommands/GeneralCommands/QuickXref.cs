using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Linq;

namespace ElectricalCommands
{
    public partial class GeneralCommands
    {
        [CommandMethod("QUICKXREF", CommandFlags.Modal)]
        [CommandMethod("QXR", CommandFlags.Modal)]
        public void QuickXref()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // Step 1: Ensure we are in Model Space
                if (!db.TileMode)
                {
                    Application.SetSystemVariable("TILEMODE", 1);
                    ed.WriteMessage("\nSwitched to Model Space.");
                }

                // Step 2: Validate drawing is saved
                string drawingPath = doc.Name;
                if (string.IsNullOrWhiteSpace(drawingPath) || !File.Exists(drawingPath))
                {
                    ed.WriteMessage("\nError: The current drawing must be saved before using QUICKXREF.");
                    return;
                }

                string drawingDir = Path.GetDirectoryName(drawingPath);

                // Step 3: Find the Xrefs folder
                string xrefsFolder = FindXrefsFolder(drawingDir);
                if (xrefsFolder == null)
                {
                    ed.WriteMessage("\nError: Could not find an 'Xrefs' folder. Searched up from: " + drawingDir);
                    return;
                }

                ed.WriteMessage("\nFound Xrefs folder: " + xrefsFolder);

                // Step 4: List DWG files
                string[] dwgFiles = Directory.GetFiles(xrefsFolder, "*.dwg");
                if (dwgFiles.Length == 0)
                {
                    ed.WriteMessage("\nNo .dwg files found in: " + xrefsFolder);
                    return;
                }

                Array.Sort(dwgFiles, StringComparer.OrdinalIgnoreCase);

                // Step 5: Show file picker dialog
                var pickerWindow = new XrefPickerWindow(dwgFiles);
                bool? dialogResult = pickerWindow.ShowDialog();

                if (dialogResult != true || string.IsNullOrEmpty(pickerWindow.SelectedFile))
                {
                    ed.WriteMessage("\nCommand cancelled.");
                    return;
                }

                string selectedFile = pickerWindow.SelectedFile;
                string blockName = Path.GetFileNameWithoutExtension(selectedFile);

                // Step 6: Get insertion point
                PromptPointOptions ppo = new PromptPointOptions("\nSelect insertion point for XREF: ");
                ppo.AllowNone = false;
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK) return;

                Point3d insertionPoint = ppr.Value;

                // Step 7: Attach XREF and create block reference
                string relativePath = GetRelativePath(drawingDir, selectedFile);

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        // Check if this XREF is already loaded
                        ObjectId xrefDefId = ObjectId.Null;
                        bool alreadyLoaded = false;

                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                            if (btr.IsFromExternalReference &&
                                string.Equals(btr.Name, blockName, StringComparison.OrdinalIgnoreCase))
                            {
                                xrefDefId = btrId;
                                alreadyLoaded = true;
                                ed.WriteMessage($"\nXREF '{blockName}' is already loaded. Creating new reference.");
                                break;
                            }
                        }

                        if (!alreadyLoaded)
                        {
                            xrefDefId = db.AttachXref(relativePath, blockName);
                        }

                        // Create the BlockReference in model space
                        BlockReference blockRef = new BlockReference(insertionPoint, xrefDefId);
                        blockRef.ScaleFactors = new Scale3d(1.0);
                        blockRef.Rotation = 0.0;
                        blockRef.Layer = "0";

                        modelSpace.AppendEntity(blockRef);
                        tr.AddNewlyCreatedDBObject(blockRef, true);

                        tr.Commit();
                        ed.WriteMessage($"\nAttached XREF: {blockName}");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nFailed to attach XREF: {ex.Message}");
                    }
                }

                ed.Regen();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred: {ex.Message}");
            }
        }

        private static string FindXrefsFolder(string startDir)
        {
            string current = startDir;

            for (int i = 0; i < 5; i++)
            {
                if (string.IsNullOrEmpty(current)) break;

                string parent = Path.GetDirectoryName(current);
                if (string.IsNullOrEmpty(parent)) break;

                // Check for "Xrefs" as a sibling folder (same parent)
                string candidate = Path.Combine(parent, "Xrefs");
                if (Directory.Exists(candidate))
                    return candidate;

                current = parent;
            }

            return null;
        }

        private static string GetRelativePath(string fromDir, string toFile)
        {
            try
            {
                if (!fromDir.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    fromDir += Path.DirectorySeparatorChar;

                Uri fromUri = new Uri(fromDir);
                Uri toUri = new Uri(toFile);

                Uri relativeUri = fromUri.MakeRelativeUri(toUri);
                string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

                return relativePath.Replace('/', Path.DirectorySeparatorChar);
            }
            catch
            {
                return toFile;
            }
        }
    }
}
