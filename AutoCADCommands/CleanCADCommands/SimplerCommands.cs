using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint BM_CLICK = 0x00F5;

        // *** MODIFICATION: Restored the missing method ***
        /// <summary>
        /// Unlocks, thaws, and turns on all layers in the drawing to ensure visibility and editability.
        /// </summary>
        internal static void EnsureAllLayersVisibleAndUnlocked(Database db, Editor ed)
        {
            ed.WriteMessage("\nEnsuring all layers are visible and unlocked...");
            int unlocked = 0;
            int thawed = 0;
            int turnedOn = 0;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId originalClayerId = db.Clayer;

                    if (lt.Has("0") && db.Clayer != lt["0"])
                    {
                        db.Clayer = lt["0"];
                    }

                    foreach (ObjectId layerId in lt)
                    {
                        if (layerId.IsErased) continue;
                        try
                        {
                            var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);

                            if (ltr.IsLocked)
                            {
                                ltr.IsLocked = false;
                                unlocked++;
                            }
                            if (ltr.IsFrozen)
                            {
                                ltr.IsFrozen = false;
                                thawed++;
                            }
                            if (ltr.IsOff)
                            {
                                ltr.IsOff = false;
                                turnedOn++;
                            }
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
                        {
                            db.Clayer = originalClayerId;
                        }
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
                                SetForegroundWindow(dlg);
                                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                            }
                            return;
                        }
                    }
                    catch { }
                    System.Threading.Thread.Sleep(150);
                }
            });
        }

        private class ImagePlacement
        {
            public string Path;
            public Point3d Pos;
            public Vector3d U;
            public Vector3d V;
            public ObjectId OriginalEntityId;
            public Point2d[] ClipBoundary;
            public ObjectId TargetBtrId = ObjectId.Null;
        }

        private static readonly Queue<ImagePlacement> _pending = new Queue<ImagePlacement>();
        private static ObjectId _lastPastedOle = ObjectId.Null;
        private static dynamic _pptAppShared = null;
        private static dynamic _pptPresentationShared = null;
        private static ObjectId _savedClayer = ObjectId.Null;
        private static readonly HashSet<ObjectId> _imageDefsToPurge = new HashSet<ObjectId>();
        internal static readonly HashSet<ObjectId> _pdfDefinitionsToDetach = new HashSet<ObjectId>();
        private static bool _chainFinalizeAfterEmbed = false;
        private static readonly HashSet<ObjectId> _xrefsToDetach = new HashSet<ObjectId>();

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
                                {
                                    _pptPresentationShared.Slides.Add(1, 12);
                                }
                                return true;
                            }
                            catch
                            {
                                _pptPresentationShared = presentations.Add();
                                _pptPresentationShared.Slides.Add(1, 12);
                                return true;
                            }
                        }
                        else
                        {
                            _pptPresentationShared = presentations.Add();
                            _pptPresentationShared.Slides.Add(1, 12);
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
                    ed.WriteMessage("\nPowerPoint is not installed.");
                    return false;
                }
                _pptAppShared = Activator.CreateInstance(pptType);
                try { _pptAppShared.Visible = true; } catch { }
                _pptPresentationShared = _pptAppShared.Presentations.Add();
                _pptPresentationShared.Slides.Add(1, 12); // ppLayoutBlank
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to start or access PowerPoint: {ex.Message}");
                ClosePowerPoint();
                return false;
            }
        }

        private static void ClosePowerPoint(Editor ed = null)
        {
            if (_pptPresentationShared != null)
            {
                try { _pptPresentationShared.Saved = -1; _pptPresentationShared.Close(); } catch { }
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

        private static bool PrepareClipboardWithCroppedImage(ImagePlacement placement, Editor ed)
        {
            try
            {
                if (!EnsurePowerPoint(ed)) return false;
                dynamic slide = _pptPresentationShared.Slides[1];
                var shapes = slide.Shapes;
                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes[i].Delete(); } catch { }
                }

                Shape pic = shapes.AddPicture(placement.Path, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10);

                if (placement.ClipBoundary != null && placement.ClipBoundary.Length >= 2)
                {
                    pic.ScaleHeight(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                    pic.ScaleWidth(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);

                    var clip = placement.ClipBoundary;
                    double minX = clip.Min(p => p.X);
                    double minY = clip.Min(p => p.Y);
                    double maxX = clip.Max(p => p.X);
                    double maxY = clip.Max(p => p.Y);

                    float picWidth = pic.Width;
                    float picHeight = pic.Height;

                    pic.PictureFormat.CropLeft = (float)(minX * picWidth);
                    pic.PictureFormat.CropTop = (float)((1.0 - maxY) * picHeight);
                    pic.PictureFormat.CropRight = (float)((1.0 - maxX) * picWidth);
                    pic.PictureFormat.CropBottom = (float)(minY * picHeight);
                }

                try { System.Windows.Forms.Clipboard.Clear(); } catch { }
                pic.Copy();

                dynamic pngRange = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG);
                if (pngRange != null)
                {
                    pngRange.Copy();
                    pngRange.Delete();
                }
                pic.Delete();
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prepare clipboard for '{placement.Path}': {ex.Message}");
                return false;
            }
        }

        private static bool PrepareClipboardWithImageShared(ImagePlacement placement, Editor ed)
        {
            placement.ClipBoundary = null;
            return PrepareClipboardWithCroppedImage(placement, ed);
        }

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
                    int purged = 0;
                    foreach (var defId in _imageDefsToPurge)
                    {
                        if (defId.IsErased) continue;
                        try
                        {
                            var def = tr.GetObject(defId, OpenMode.ForWrite);
                            def.Erase();
                            purged++;
                        }
                        catch { }
                    }
                    if (purged > 0) ed.WriteMessage($"\nPurged {purged} unused image definition(s).");
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
                    int detachedCount = 0;
                    var keysToRemove = new List<string>();

                    foreach (DBDictionaryEntry entry in pdfDict)
                    {
                        if (_pdfDefinitionsToDetach.Contains(entry.Value))
                        {
                            keysToRemove.Add(entry.Key);
                        }
                    }

                    foreach (string key in keysToRemove)
                    {
                        try
                        {
                            var defId = pdfDict.GetAt(key);
                            pdfDict.Remove(key);
                            var defObj = tr.GetObject(defId, OpenMode.ForWrite);
                            defObj.Erase();
                            detachedCount++;
                        }
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

        private static void Db_ObjectAppended(object sender, ObjectEventArgs e)
        {
            if (e.DBObject is Ole2Frame)
            {
                _lastPastedOle = e.DBObject.ObjectId;
            }
        }

        private static void DetachXrefs(Database db, Editor ed)
        {
            if (_xrefsToDetach.Count == 0) return;
            ed.WriteMessage($"\nAttempting to detach {_xrefsToDetach.Count} XREF(s) ...");
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

        private static bool EnsurePlacementSpace(Document doc, ImagePlacement placement, Editor ed)
        {
            if (placement.TargetBtrId.IsNull || doc.Database.CurrentSpaceId == placement.TargetBtrId)
            {
                return true;
            }

            try
            {
                string layoutName;
                using (var tr = doc.Database.TransactionManager.StartOpenCloseTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(placement.TargetBtrId, OpenMode.ForRead);
                    var layout = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);
                    layoutName = layout.LayoutName;
                }
                if (!string.Equals(LayoutManager.Current.CurrentLayout, layoutName, StringComparison.OrdinalIgnoreCase))
                {
                    LayoutManager.Current.CurrentLayout = layoutName;
                }
                return doc.Database.CurrentSpaceId == placement.TargetBtrId;
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\nFailed to switch to target layout: {ex.Message}");
                return false;
            }
        }

        private static void ProcessPastedOle(ObjectId oleId, ImagePlacement target, Database db, Editor ed)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ole = tr.GetObject(oleId, OpenMode.ForWrite) as Ole2Frame;
                if (ole == null)
                {
                    ed.WriteMessage("\nError: Could not find the pasted OLE object for processing.");
                    tr.Commit();
                    return;
                }

                try
                {
                    var ext = ole.GeometricExtents;
                    double oleWidth = Math.Max(1e-8, ext.MaxPoint.X - ext.MinPoint.X);
                    double oleHeight = Math.Max(1e-8, ext.MaxPoint.Y - ext.MinPoint.Y);
                    double destWidth = target.U.Length;
                    double destHeight = target.V.Length;

                    if (destWidth < 1e-8 || destHeight < 1e-8)
                    {
                        ed.WriteMessage($"\nSkipping transform due to zero-sized destination for {target.OriginalEntityId}.");
                        ole.Erase();
                    }
                    else
                    {
                        var align = Matrix3d.AlignCoordinateSystem(
                            ext.MinPoint, new Vector3d(oleWidth, 0, 0), new Vector3d(0, oleHeight, 0), Vector3d.ZAxis,
                            target.Pos, target.U, target.V, target.U.CrossProduct(target.V)
                        );
                        ole.TransformBy(align);
                        try { ole.Layer = "0"; } catch { }

                        var btr = (BlockTableRecord)tr.GetObject(ole.OwnerId, OpenMode.ForWrite);
                        var dot = (DrawOrderTable)tr.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite);
                        dot.MoveAbove(new ObjectIdCollection { ole.ObjectId }, target.OriginalEntityId);
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nFailed to transform pasted OLE: {ex.Message}");
                }

                try
                {
                    var originalEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForWrite) as Entity;
                    if (originalEnt is RasterImage imgEnt && !imgEnt.ImageDefId.IsNull)
                    {
                        _imageDefsToPurge.Add(imgEnt.ImageDefId);
                    }
                    originalEnt?.Erase();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: failed to erase source entity: {ex.Message}");
                }

                tr.Commit();
            }
        }

        public static bool TryGetTitleBlockOutlinePointsForEmbed(Database db, out Point3d[] poly)
        {
            poly = null;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lm = LayoutManager.Current;
                    if (lm == null) return false;

                    var layId = lm.GetLayoutId(lm.CurrentLayout);
                    var layout = (Layout)tr.GetObject(layId, OpenMode.ForRead);

                    var psId = layout.BlockTableRecordId;
                    var psBtr = (BlockTableRecord)tr.GetObject(psId, OpenMode.ForRead);

                    string[] hints = { "x-tb", "title", "tblock", "border", "sheet" };
                    BlockReference best = null;
                    double bestArea = 0;

                    foreach (ObjectId id in psBtr)
                    {
                        var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (br == null) continue;

                        string name = null;
                        try
                        {
                            var btr = br.IsDynamicBlock
                                ? (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)
                                : (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            name = btr?.Name;
                        }
                        catch { }

                        if (string.IsNullOrEmpty(name)) continue;

                        bool matches = hints.Any(h => name.IndexOf(h, StringComparison.OrdinalIgnoreCase) >= 0);
                        if (!matches) continue;

                        try
                        {
                            var ext = br.GeometricExtents;
                            double area = Math.Abs((ext.MaxPoint.X - ext.MinPoint.X) * (ext.MaxPoint.Y - ext.MinPoint.Y));
                            if (area > bestArea)
                            {
                                best = br;
                                bestArea = area;
                            }
                        }
                        catch { continue; }
                    }

                    if (best != null)
                    {
                        var ex = best.GeometricExtents;
                        poly = new[] {
                            new Point3d(ex.MinPoint.X, ex.MinPoint.Y, 0), new Point3d(ex.MaxPoint.X, ex.MinPoint.Y, 0),
                            new Point3d(ex.MaxPoint.X, ex.MaxPoint.Y, 0), new Point3d(ex.MinPoint.X, ex.MaxPoint.Y, 0)
                        };
                        return true;
                    }

                    poly = new[] {
                        new Point3d(db.Pextmin.X, db.Pextmin.Y, 0), new Point3d(db.Pextmax.X, db.Pextmin.Y, 0),
                        new Point3d(db.Pextmax.X, db.Pextmax.Y, 0), new Point3d(db.Pextmin.X, db.Pextmax.Y, 0)
                    };
                    return true;
                }
            }
            catch { return false; }
        }

        public static void ZoomToTitleBlockForEmbed(Editor ed, Point3d[] poly)
        {
            if (ed == null || poly == null || poly.Length < 2) return;
            try
            {
                var ext = new Extents3d(poly[0], poly[1]);
                for (int i = 2; i < poly.Length; i++) ext.AddPoint(poly[i]);

                double margin = Math.Max(ext.MaxPoint.X - ext.MinPoint.X, ext.MaxPoint.Y - ext.MinPoint.Y) * 0.05;
                Point3d pMin = new Point3d(ext.MinPoint.X - margin, ext.MinPoint.Y - margin, 0);
                Point3d pMax = new Point3d(ext.MaxPoint.X + margin, ext.MaxPoint.Y + margin, 0);

                using (var view = ed.GetCurrentView())
                {
                    view.Width = pMax.X - pMin.X;
                    view.Height = pMax.Y - pMin.Y;
                    view.CenterPoint = new Point2d((pMin.X + pMax.X) / 2.0, (pMin.Y + pMax.Y) / 2.0);
                    ed.SetCurrentView(view);
                }
                ed.Regen();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during zoom: {ex.Message}");
            }
        }
    }
}