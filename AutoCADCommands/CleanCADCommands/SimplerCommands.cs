using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Globalization;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Diagnostics;

// Required for in-memory image rotation
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        internal enum PlacementSource
        {
            Unknown = 0,
            Pdf = 1,
            Xref = 2
        }

        // Win32 helpers to auto-dismiss the "OLE Text Size" dialog
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint BM_CLICK = 0x00F5;
        private const uint WM_COMMAND = 0x0111;
        private const int IDOK = 1;

        // Shared: unlock/thaw layers
        internal static void EnsureAllLayersVisibleAndUnlocked(Database db, Editor ed)
        {
            ed.WriteMessage("\nEnsuring all layers are visible and unlocked...");
            int unlocked = 0, thawed = 0, turnedOn = 0;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId originalClayerId = db.Clayer;

                    if (lt.Has("0") && db.Clayer != lt["0"])
                        db.Clayer = lt["0"];

                    foreach (ObjectId layerId in lt)
                    {
                        if (layerId.IsErased) continue;
                        try
                        {
                            var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                            if (ltr.IsLocked) { ltr.IsLocked = false; unlocked++; }
                            if (ltr.IsFrozen) { ltr.IsFrozen = false; thawed++; }
                            if (ltr.IsOff) { ltr.IsOff = false; turnedOn++; }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nCould not modify layer {layerId}: {ex.Message}");
                        }
                    }

                    try
                    {
                        var originalLtr = (LayerTableRecord)tr.GetObject(originalClayerId, OpenMode.ForRead);
                        if (!originalLtr.IsErased && !originalLtr.IsFrozen)
                            db.Clayer = originalClayerId;
                    }
                    catch { }

                    tr.Commit();
                }
                ed.WriteMessage($"\nUnlocked: {unlocked}, Thawed: {thawed}, Turned On: {turnedOn} layer(s).");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred while modifying layers: {ex.Message}");
            }
        }

        private static void StartOleTextSizeDialogCloser(double seconds = 120)
        {
            Task.Run(() =>
            {
                var until = DateTime.UtcNow.AddSeconds(seconds);
                while (DateTime.UtcNow < until)
                {
                    try
                    {
                        IntPtr dlg = FindWindow("#32770", "OLE Text Size");
                        if (dlg != IntPtr.Zero)
                        {
                            IntPtr ok = FindWindowEx(dlg, IntPtr.Zero, "Button", "OK");
                            if (ok != IntPtr.Zero)
                            {
                                SendMessage(ok, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                            }
                            else
                            {
                                SendMessage(dlg, WM_COMMAND, (IntPtr)IDOK, IntPtr.Zero);
                                SetForegroundWindow(dlg);
                                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                            }
                            return;
                        }
                    }
                    catch { }
                    Thread.Sleep(150);
                }
            });
        }

        // Placement description used by both commands, distinguished by Source
        private class ImagePlacement
        {
            public string Path;
            public Point3d Pos;
            public Vector3d U;
            public Vector3d V;
            public ObjectId OriginalEntityId;
            public ObjectId TargetBtrId = ObjectId.Null;
            public Point2d[] ClipBoundary; // PDFs only
            public PlacementSource Source = PlacementSource.Unknown;
        }

        private static readonly Queue<ImagePlacement> _pending = new Queue<ImagePlacement>();
        private static ObjectId _lastPastedOle = ObjectId.Null;
        private static bool _handlersAttached = false;
        private static dynamic _pptAppShared = null;
        private static dynamic _pptPresentationShared = null;
        private static ObjectId _savedClayer = ObjectId.Null;
        private static readonly HashSet<ObjectId> _imageDefsToPurge = new HashSet<ObjectId>();
        internal static readonly HashSet<ObjectId> _pdfDefinitionsToDetach = new HashSet<ObjectId>();
        private static bool _chainFinalizeAfterEmbed = false;
        private static ImagePlacement _activePlacement = null;
        private static Document _activePasteDocument = null;
        private static bool _waitingForPasteStart = false;
        private static bool _pastePointHandlerAttached = false;
        private static bool _isEmbeddingProcessActive = false;
        private static bool _isCleanSheetWorkflowActive = false;
        internal static bool _skipLayerFreezing = false;
        private static ObjectId _finalPastedOleForZoom = ObjectId.Null;
        private static readonly HashSet<ObjectId> _xrefsToDetach = new HashSet<ObjectId>();

        // Shared PowerPoint lifecycle
        private static bool EnsurePowerPoint(Editor ed)
        {
            try
            {
                if (_pptAppShared != null)
                {
                    try
                    {
                        var presentations = _pptAppShared.Presentations;
                        try { _pptAppShared.Visible = true; } catch { }

                        if (_pptPresentationShared != null)
                        {
                            try
                            {
                                if (_pptPresentationShared.Slides.Count < 1)
                                    _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                                return true;
                            }
                            catch
                            {
                                _pptPresentationShared = presentations.Add(MsoTriState.msoFalse);
                                _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                                return true;
                            }
                        }
                        else
                        {
                            _pptPresentationShared = presentations.Add(MsoTriState.msoFalse);
                            _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                            return true;
                        }
                    }
                    catch
                    {
                        try { Marshal.FinalReleaseComObject(_pptAppShared); } catch { }
                        _pptAppShared = null;
                        _pptPresentationShared = null;
                    }
                }

                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    ed.WriteMessage("\nPowerPoint is not installed (COM ProgID not found).");
                    return false;
                }
                _pptAppShared = Activator.CreateInstance(pptType);
                try { _pptAppShared.Visible = true; } catch { }
                var presNew = _pptAppShared.Presentations;
                _pptPresentationShared = presNew.Add(MsoTriState.msoFalse);
                _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to start or access PowerPoint: {ex.Message}");
                ClosePowerPoint(ed);
                return false;
            }
        }

        private static void ClosePowerPoint(Editor ed = null)
        {
            try
            {
                if (_pptPresentationShared != null)
                {
                    try { _pptPresentationShared.Saved = MsoTriState.msoTrue; } catch { }
                    try { _pptPresentationShared.Close(); } catch { }
                    try { Marshal.FinalReleaseComObject(_pptPresentationShared); } catch { }
                    _pptPresentationShared = null;
                }
                if (_pptAppShared != null)
                {
                    try { _pptAppShared.Quit(); } catch { }
                    try { Marshal.FinalReleaseComObject(_pptAppShared); } catch { }
                    _pptAppShared = null;
                }
            }
            catch (System.Exception ex)
            {
                try { ed?.WriteMessage($"\nWarning: failed to close PowerPoint: {ex.Message}"); } catch { }
                _pptPresentationShared = null;
                _pptAppShared = null;
            }
        }

        // Shared: put given image on clipboard via PPT
        private static bool PrepareClipboardWithImageShared(ImagePlacement placement, Editor ed)
        {
            dynamic slide = null;
            dynamic shapes = null;
            Shape pic = null;

            try
            {
                if (!EnsurePowerPoint(ed)) return false;
                slide = _pptPresentationShared.Slides[1];
                shapes = slide.Shapes;

                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes[i].Delete(); } catch { }
                }

                pic = shapes.AddPicture(placement.Path, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10);
                pic.Copy();
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prepare clipboard for '{placement.Path}': {ex.Message}");
                return false;
            }
            finally
            {
                if (pic != null) Marshal.ReleaseComObject(pic);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (slide != null) Marshal.ReleaseComObject(slide);
            }
        }

        // Shared: resolve image/PDF paths
        private static string ResolveImagePath(Database db, string rawPath)
        {
            if (string.IsNullOrWhiteSpace(rawPath)) return null;
            string p = rawPath.Trim().Trim('"').Replace('/', '\\');
            if (Path.IsPathRooted(p) && File.Exists(p)) return p;

            try
            {
                string dbFile = db?.Filename;
                if (!string.IsNullOrWhiteSpace(dbFile))
                {
                    string baseDir = Path.GetDirectoryName(dbFile);
                    if (!string.IsNullOrWhiteSpace(baseDir))
                    {
                        if (Path.IsPathRooted(p) && p.StartsWith("\\") && !p.StartsWith("\\\\"))
                        {
                            string drive = Path.GetPathRoot(baseDir);
                            string combined = Path.Combine(drive ?? string.Empty, p.TrimStart('\\'));
                            string full = Path.GetFullPath(combined);
                            if (File.Exists(full)) return full;
                        }

                        string candidate = Path.GetFullPath(Path.Combine(baseDir, p));
                        if (File.Exists(candidate)) return candidate;
                    }
                }
            }
            catch { }

            try
            {
                string nameOnly = Path.GetFileName(p);
                if (!string.IsNullOrEmpty(nameOnly))
                {
                    string found = HostApplicationServices.Current.FindFile(nameOnly, db, FindFileHint.Default);
                    if (!string.IsNullOrWhiteSpace(found) && File.Exists(found)) return found;
                }
            }
            catch { }

            return null;
        }

        // Shared: hook/unhook event handlers
        private static void AttachHandlers(Database db, Document doc)
        {
            if (_handlersAttached) return;
            db.ObjectAppended += Db_ObjectAppended;
            doc.CommandWillStart += Doc_CommandWillStart;
            doc.CommandEnded += Doc_CommandEnded;
            _handlersAttached = true;
        }

        private static void DetachHandlers(Database db, Document doc)
        {
            if (!_handlersAttached) return;
            try { db.ObjectAppended -= Db_ObjectAppended; } catch { }
            try { doc.CommandWillStart -= Doc_CommandWillStart; } catch { }
            try { doc.CommandEnded -= Doc_CommandEnded; } catch { }
            if (_pastePointHandlerAttached)
            {
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint; } catch { }
                _pastePointHandlerAttached = false;
            }
            _activePlacement = null;
            _activePasteDocument = null;
            _waitingForPasteStart = false;
            _handlersAttached = false;
        }

        // Shared: purge image defs whose RasterImages were embedded
        private static void PurgeEmbeddedImageDefs(Database db, Editor ed)
        {
            if (_imageDefsToPurge.Count == 0) return;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var named = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!named.Contains("ACAD_IMAGE_DICT")) return;

                    var imageDict = (DBDictionary)tr.GetObject(named.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);
                    var keysToRemove = new List<string>();

                    foreach (DBDictionaryEntry entry in imageDict)
                    {
                        if (_imageDefsToPurge.Contains(entry.Value))
                            keysToRemove.Add(entry.Key);
                    }

                    foreach (string key in keysToRemove)
                    {
                        try { imageDict.Remove(key); } catch { }
                    }

                    int purged = 0;
                    foreach (var defId in _imageDefsToPurge)
                    {
                        try
                        {
                            var defObj = tr.GetObject(defId, OpenMode.ForWrite, false);
                            if (defObj != null && !defObj.IsErased)
                            {
                                defObj.Erase();
                                purged++;
                            }
                        }
                        catch { }
                    }

                    if (purged > 0)
                        ed.WriteMessage($"\nPurged {purged} unused image definition(s).");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to purge image definitions: {ex.Message}");
            }
            finally
            {
                _imageDefsToPurge.Clear();
            }
        }

        // Shared: detach PDF definitions recorded by PDF pipeline
        private static void DetachPdfDefinitions(Database db, Editor ed)
        {
            if (_pdfDefinitionsToDetach.Count == 0) return;
            ed.WriteMessage($"\nAttempting to detach {_pdfDefinitionsToDetach.Count} PDF definition(s)...");
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var named = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!named.Contains("ACAD_PDFDEFINITIONS")) return;

                    var pdfDict = (DBDictionary)tr.GetObject(named.GetAt("ACAD_PDFDEFINITIONS"), OpenMode.ForWrite);
                    var entriesToRemove = new List<string>();

                    foreach (DBDictionaryEntry entry in pdfDict)
                    {
                        if (_pdfDefinitionsToDetach.Contains(entry.Value))
                        {
                            entriesToRemove.Add(entry.Key);
                            try
                            {
                                var defObj = tr.GetObject(entry.Value, OpenMode.ForWrite);
                                if (!defObj.IsErased) defObj.Erase();
                            }
                            catch { }
                        }
                    }

                    int detachedCount = 0;
                    foreach (string key in entriesToRemove)
                    {
                        try { pdfDict.Remove(key); detachedCount++; }
                        catch { }
                    }

                    ed.WriteMessage($"\nSuccessfully detached {detachedCount} PDF definition(s).");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during PDF detachment: {ex.Message}");
            }
            finally
            {
                _pdfDefinitionsToDetach.Clear();
            }
        }

        // Event handlers: shared entry, but behavior split by PlacementSource

        private static void Db_ObjectAppended(object sender, ObjectEventArgs e)
        {
            if (e.DBObject is Ole2Frame)
            {
                _lastPastedOle = e.DBObject.ObjectId;
            }
        }

        private static void Doc_CommandWillStart(object sender, CommandEventArgs e)
        {
            if (!string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;
            if (!_waitingForPasteStart || _activePlacement == null || _pastePointHandlerAttached) return;
            _activePasteDocument ??= AutoCADApp.DocumentManager.MdiActiveDocument;
            AutoCADApp.Idle += Application_OnIdleSendPastePoint;
            _pastePointHandlerAttached = true;
        }

        private static void Application_OnIdleSendPastePoint(object sender, EventArgs e)
        {
            try
            {
                AutoCADApp.Idle -= Application_OnIdleSendPastePoint;
                _pastePointHandlerAttached = false;
                var doc = _activePasteDocument ?? AutoCADApp.DocumentManager.MdiActiveDocument;
                var placement = _activePlacement;
                if (doc == null || placement == null) return;
                if (!string.Equals(doc.CommandInProgress, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;

                string x = placement.Pos.X.ToString("G17", CultureInfo.InvariantCulture);
                string y = placement.Pos.Y.ToString("G17", CultureInfo.InvariantCulture);
                doc.SendStringToExecute($"{x},{y}\n", true, false, false);
            }
            finally
            {
                _waitingForPasteStart = false;
            }
        }

        private static void Doc_CommandEnded(object sender, CommandEventArgs e)
        {
            if (!_isEmbeddingProcessActive ||
                !string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase))
                return;

            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            _waitingForPasteStart = false;
            if (_pastePointHandlerAttached)
            {
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint; } catch { }
                _pastePointHandlerAttached = false;
            }

            if (_pending.Count == 0 && _activePlacement == null)
                return;

            if (_lastPastedOle.IsNull)
            {
                ed.WriteMessage("\nSkipping an item because no OLE object was created.");
                if (_pending.Count > 0) _pending.Dequeue();
                _activePlacement = null;

                if (_pending.Count > 0)
                    ProcessNextPaste(doc, ed);
                else
                    FinishEmbeddingRun(doc, ed, db);



                return;
            }

            var target = _pending.Count > 0 ? _pending.Dequeue() : _activePlacement;
            if (target == null)
            {
                ed.WriteMessage("\nError: Could not retrieve active image placement. Aborting.");
                FinishEmbeddingRun(doc, ed, db);
                return;
            }
            _activePlacement = null;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ole = tr.GetObject(_lastPastedOle, OpenMode.ForWrite) as Ole2Frame;

                if (ole != null)
                {
                    try
                    {
                        if (target.Source == PlacementSource.Pdf)
                        {
                            ApplyPdfPlacementTransform(tr, ed, ole, target);
                        }
                        else // Xref or Unknown -> treat as Xref-style raster
                        {
                            ApplyXrefPlacementTransform(tr, ed, ole, target);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nFailed to transform pasted OLE: {ex.Message}");
                    }
                }

                // Erase original entity and record defs if applicable
                try
                {
                    var originalEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForWrite, false) as Entity;
                    if (originalEnt != null)
                    {
                        if (originalEnt is RasterImage imgEnt && !imgEnt.ImageDefId.IsNull)
                            _imageDefsToPurge.Add(imgEnt.ImageDefId);

                        originalEnt.Erase();
                    }
                }
                catch { }

                tr.Commit();
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Thread.Sleep(50);

            if (_pending.Count > 0)
            {
                ProcessNextPaste(doc, ed);
            }
            else
            {
                _finalPastedOleForZoom = _lastPastedOle;
                FinishEmbeddingRun(doc, ed, db);
            }

            _lastPastedOle = ObjectId.Null;
        }

        // PDF-specific final placement (from your working EMBEDFROMPDFS implementation)
        private static void ApplyPdfPlacementTransform(
            Transaction tr,
            Editor ed,
            Ole2Frame ole,
            ImagePlacement target)
        {
            var ext = ole.GeometricExtents;
            Point3d oleOrigin = ext.MinPoint;
            double oleWidth = Math.Max(1e-8, ext.MaxPoint.X - ext.MinPoint.X);
            double oleHeight = Math.Max(1e-8, ext.MaxPoint.Y - ext.MinPoint.Y);

            Vector3d destU = target.U;
            Vector3d destV = target.V;
            double destWidth = destU.Length;
            double destHeight = destV.Length;

            if (destWidth < 1e-8 || destHeight < 1e-8)
            {
                ed.WriteMessage("\nSkipping PDF transform: invalid target geometry.");
                ole.Erase();
                return;
            }

            double oleAspect = oleWidth / oleHeight;
            double destAspect = destWidth / destHeight;

            Vector3d finalU = destU;
            Vector3d finalV = destV;
            Point3d finalPos = target.Pos;

            if (Math.Abs(oleAspect - destAspect) > 1e-6)
            {
                if (oleAspect > destAspect)
                {
                    double newHeight = destWidth / oleAspect;
                    finalV = destV.GetNormal() * newHeight;
                    finalPos = target.Pos + (destV - finalV) / 2.0;
                }
                else
                {
                    double newWidth = destHeight * oleAspect;
                    finalU = destU.GetNormal() * newWidth;
                    finalPos = target.Pos + (destU - finalU) / 2.0;
                }
            }

            Vector3d destNormal = finalU.CrossProduct(finalV);
            if (destNormal.Length < 1e-8)
                destNormal = Vector3d.ZAxis;

            Vector3d oleU = new Vector3d(oleWidth, 0.0, 0.0);
            Vector3d oleV = new Vector3d(0.0, oleHeight, 0.0);
            Vector3d oleNormal = oleU.CrossProduct(oleV);

            Matrix3d align = Matrix3d.AlignCoordinateSystem(
                oleOrigin,
                oleU,
                oleV,
                oleNormal,
                finalPos,
                finalU,
                finalV,
                destNormal);

            ole.TransformBy(align);

            try { ole.Layer = "0"; } catch { }

            try
            {
                if (!target.OriginalEntityId.IsNull)
                {
                    var origEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForRead, false) as Entity;
                    if (origEnt != null && !origEnt.IsErased)
                    {
                        var btr = tr.GetObject(ole.OwnerId, OpenMode.ForWrite) as BlockTableRecord;
                        if (btr != null && btr.DrawOrderTableId.IsValid && origEnt.OwnerId == btr.ObjectId)
                        {
                            var dot = tr.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                            if (dot != null)
                            {
                                var ids = new ObjectIdCollection(new[] { ole.ObjectId });
                                dot.MoveAbove(ids, origEnt.ObjectId);
                            }
                        }
                    }
                }
            }
            catch (System.Exception exOrder)
            {
                ed.WriteMessage($"\nWarning: Could not match draw order (PDF): {exOrder.Message}");
            }
        }

        // XREF-specific final placement: matches your working rotated-image behavior
        private static void ApplyXrefPlacementTransform(
            Transaction tr,
            Editor ed,
            Ole2Frame ole,
            ImagePlacement target)
        {
            var ext = ole.GeometricExtents;
            Point3d oleOrigin = ext.MinPoint;
            double oleWidth = Math.Max(1e-8, ext.MaxPoint.X - ext.MinPoint.X);
            double oleHeight = Math.Max(1e-8, ext.MaxPoint.Y - ext.MinPoint.Y);

            Vector3d destU = target.U;
            Vector3d destV = target.V;
            double destWidth = destU.Length;
            double destHeight = destV.Length;

            if (destWidth < 1e-8 || destHeight < 1e-8)
            {
                ed.WriteMessage("\nSkipping XREF transform: invalid target geometry.");
                ole.Erase();
                return;
            }

            double oleAspect = oleWidth / oleHeight;
            double destAspect = destWidth / destHeight;

            Vector3d finalU = destU;
            Vector3d finalV = destV;
            Point3d finalPos = target.Pos;

            if (Math.Abs(oleAspect - destAspect) > 1e-6)
            {
                if (oleAspect > destAspect)
                {
                    double newHeight = destWidth / oleAspect;
                    finalV = destV.GetNormal() * newHeight;
                    finalPos = target.Pos + (destV - finalV) / 2.0;
                }
                else
                {
                    double newWidth = destHeight * oleAspect;
                    finalU = destU.GetNormal() * newWidth;
                    finalPos = target.Pos + (destU - finalU) / 2.0;
                }
            }

            Vector3d destNormal = finalU.CrossProduct(finalV);
            if (destNormal.Length < 1e-8)
                destNormal = Vector3d.ZAxis;

            Vector3d oleU = new Vector3d(oleWidth, 0.0, 0.0);
            Vector3d oleV = new Vector3d(0.0, oleHeight, 0.0);
            Vector3d oleNormal = oleU.CrossProduct(oleV);

            Matrix3d align = Matrix3d.AlignCoordinateSystem(
                oleOrigin,
                oleU,
                oleV,
                oleNormal,
                finalPos,
                finalU,
                finalV,
                destNormal);

            ole.TransformBy(align);

            try { ole.Layer = "0"; } catch { }

            try
            {
                if (!target.OriginalEntityId.IsNull)
                {
                    var origEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForRead, false) as Entity;
                    if (origEnt != null && !origEnt.IsErased)
                    {
                        var btr = tr.GetObject(ole.OwnerId, OpenMode.ForWrite) as BlockTableRecord;
                        if (btr != null && btr.DrawOrderTableId.IsValid && origEnt.OwnerId == btr.ObjectId)
                        {
                            var dot = tr.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                            if (dot != null)
                            {
                                var ids = new ObjectIdCollection(new[] { ole.ObjectId });
                                dot.MoveAbove(ids, origEnt.ObjectId);
                            }
                        }
                    }
                }
            }
            catch (System.Exception exOrder)
            {
                ed.WriteMessage($"\nWarning: Could not match draw order (XREF): {exOrder.Message}");
            }
        }

        // Final cleanup (shared)
        private static void FinalCleanupOnIdle(object sender, EventArgs e)
        {
            AutoCADApp.Idle -= FinalCleanupOnIdle;
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null || !_isEmbeddingProcessActive) return;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nPerforming final cleanup for embedding process...");
                _isEmbeddingProcessActive = false;
                _finalPastedOleForZoom = ObjectId.Null;

                DetachHandlers(db, doc);
                DetachXrefs(db, ed);
                try { doc.SendStringToExecute("_.REGEN ", false, false, true); } catch { }

                PurgeEmbeddedImageDefs(db, ed);
                DetachPdfDefinitions(db, ed);

                WindowOrchestrator.EndPptInteraction();
                ClosePowerPoint(ed);

                try
                {
                    if (!_savedClayer.IsNull) db.Clayer = _savedClayer;
                }
                catch { }
                _savedClayer = ObjectId.Null;
                _skipLayerFreezing = false;

                if (_chainFinalizeAfterEmbed)
                {
                    _chainFinalizeAfterEmbed = false;
                    ed.WriteMessage("\nEMBEDFROMXREFS complete. Chaining FINALIZE and DETACHREMAININGXREFS...");
                    doc.SendStringToExecute("_.FINALIZE _.DETACHREMAININGXREFS ", true, false, false);
                }
                else if (_isCleanSheetWorkflowActive)
                {
                    _isCleanSheetWorkflowActive = false;
                    ed.WriteMessage("\nEMBEDFROMXREFS complete. Chaining next commands...");
                    doc.SendStringToExecute("_.EMBEDFROMPDFS _.CLEANPS _.VP2PL _.FINALIZE _.DETACHREMAININGXREFS _.ZOOMTOLASTTB ", true, false, false);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during final cleanup: {ex.Message}");
                _isCleanSheetWorkflowActive = false;
                _chainFinalizeAfterEmbed = false;
                _skipLayerFreezing = false;
            }
        }

        // XREF detachment helper
        private static void DetachXrefs(Database db, Editor ed)
        {
            if (_xrefsToDetach.Count == 0) return;
            ed.WriteMessage($"\nAttempting to detach {_xrefsToDetach.Count} XREF(s)...");
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId xrefId in _xrefsToDetach)
                {
                    try
                    {
                        var btr = tr.GetObject(xrefId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null && btr.IsFromExternalReference && btr.IsResolved)
                        {
                            db.DetachXref(xrefId);
                            ed.WriteMessage($"\nSuccessfully detached XREF: {btr.Name}");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nFailed to detach XREF {xrefId}: {ex.Message}");
                    }
                }
                tr.Commit();
            }
            _xrefsToDetach.Clear();
        }

        // Shared: finalize run by scheduling idle cleanup
        private static void FinishEmbeddingRun(Document doc, Editor ed, Database db)
        {
            AutoCADApp.Idle += FinalCleanupOnIdle;
        }

        // Shared: map target BTR to layout name
        private static string GetLayoutNameFromBtrId(Database db, ObjectId btrId)
        {
            if (btrId.IsNull) return "Model";
            using (var tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr != null && btr.IsLayout)
                {
                    var layout = tr.GetObject(btr.LayoutId, OpenMode.ForRead) as Layout;
                    return layout?.LayoutName ?? "Model";
                }
                return "Model";
            }
        }

        // Shared: start next paste, but actual transform logic depends on Source in Doc_CommandEnded
        private static void ProcessNextPaste(Document doc, Editor ed)
        {
            if (!_isEmbeddingProcessActive || _pending.Count == 0)
            {
                if (_isEmbeddingProcessActive)
                    FinishEmbeddingRun(doc, ed, doc.Database);
                return;
            }

            var target = _pending.Peek();
            var db = doc.Database;
            string targetLayoutName = GetLayoutNameFromBtrId(db, target.TargetBtrId);
            if (string.IsNullOrEmpty(targetLayoutName))
            {
                _pending.Dequeue();
                ProcessNextPaste(doc, ed);
                return;
            }

            if (!PrepareClipboardWithImageShared(target, ed))
            {
                _pending.Dequeue();
                ProcessNextPaste(doc, ed);
                return;
            }

            StartOleTextSizeDialogCloser(120);
            _activePlacement = target;
            _activePasteDocument = doc;
            _waitingForPasteStart = true;

            if (_pastePointHandlerAttached)
            {
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint; } catch { }
                _pastePointHandlerAttached = false;
            }

            string commandString = $"_.-LAYOUT S \"{targetLayoutName}\"\n_.PASTECLIP\n";
            ed.WriteMessage($"\nActivating layout '{targetLayoutName}' for pasting...");
            doc.SendStringToExecute(commandString, true, false, false);
        }

        // Rotation helpers for EMBEDFROMXREFS

        private static Bitmap RotateImage(System.Drawing.Image image, float angle)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));

            const double toRad = Math.PI / 180.0;
            double angleRad = angle * toRad;
            double cos = Math.Abs(Math.Cos(angleRad));
            double sin = Math.Abs(Math.Sin(angleRad));
            int newWidth = (int)Math.Round(image.Width * cos + image.Height * sin);
            int newHeight = (int)Math.Round(image.Width * sin + image.Height * cos);

            var rotatedBmp = new Bitmap(newWidth, newHeight);
            rotatedBmp.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (Graphics g = Graphics.FromImage(rotatedBmp))
            {
                g.Clear(Color.Transparent);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                g.TranslateTransform(newWidth / 2f, newHeight / 2f);
                g.RotateTransform(angle);
                g.TranslateTransform(-image.Width / 2f, -image.Height / 2f);
                g.DrawImage(image, new System.Drawing.Point(0, 0));
            }

            return rotatedBmp;
        }

        internal static string PreflightRasterForPpt(string srcPath, double rotationInDegrees)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "embed");
            Directory.CreateDirectory(tempDir);
            string outPath = Path.Combine(
                tempDir,
                Path.GetFileNameWithoutExtension(srcPath) + "_ppt_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png"
            );

            using (var origImage = System.Drawing.Image.FromFile(srcPath))
            {
                try
                {
                    var fd = new FrameDimension(origImage.FrameDimensionsList[0]);
                    if (origImage.GetFrameCount(fd) > 1)
                        origImage.SelectActiveFrame(fd, 0);
                }
                catch { }

                System.Drawing.Image imageToProcess = origImage;
                Bitmap rotatedBitmap = null;

                if (Math.Abs(rotationInDegrees) > 0.1)
                {
                    rotatedBitmap = RotateImage(origImage, (float)rotationInDegrees);
                    imageToProcess = rotatedBitmap;
                }

                int maxSide = 8000;
                int w = imageToProcess.Width;
                int h = imageToProcess.Height;
                double scale = 1.0;
                if (Math.Max(w, h) > maxSide)
                    scale = (double)maxSide / Math.Max(w, h);

                int targetW = Math.Max(1, (int)Math.Round(w * scale));
                int targetH = Math.Max(1, (int)Math.Round(h * scale));

                using (var finalBmp = new Bitmap(targetW, targetH, PixelFormat.Format32bppArgb))
                using (var g = Graphics.FromImage(finalBmp))
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(imageToProcess, new Rectangle(0, 0, targetW, targetH));
                    finalBmp.Save(outPath, ImageFormat.Png);
                }

                rotatedBitmap?.Dispose();
            }

            return outPath;
        }
    }
}