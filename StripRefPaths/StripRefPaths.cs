// File: StripRefPathsCmd.cs (VERBOSE)
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;

namespace StripRefPaths
{
    public class Startup : IExtensionApplication
    {
        public void Initialize()
        {
            Application.DocumentManager.MdiActiveDocument?
                .Editor?.WriteMessage("\n[StripRefPaths] Loaded. Run STRIPREFPATHS.");
        }
        public void Terminate() { }
    }

    public class StripRefPathsCmd
    {
        // delete placeholders after commit
        private const bool DeletePlaceholdersAfterCommit = true;

        [CommandMethod("STRIPREFPATHS")]
        public static void StripRefPaths()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed  = doc.Editor;
            var db  = doc.Database;

            var dwgPath = db.Filename;
            var dwgDir  = SafeDirName(dwgPath);
            ed?.WriteMessage($"\n[INFO] DWG: {dwgPath}");
            ed?.WriteMessage($"\n[INFO] DWG folder: {dwgDir}");

            int imgCount = 0, pdfCount = 0, dgnCount = 0, dwfCount = 0, xrefUpdated = 0;
            int imgSeen = 0, pdfSeen = 0, dgnSeen = 0, dwfSeen = 0, xrefSeen = 0;
            int imgSkipped = 0, pdfSkipped = 0, dgnSkipped = 0, dwfSkipped = 0;

            var createdPlaceholders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    bool CanResolve(string fileOnly)
                    {
                        if (string.IsNullOrWhiteSpace(fileOnly)) return false;
                        try
                        {
                            var found = HostApplicationServices.Current.FindFile(fileOnly, db, FindFileHint.Default);
                            return !string.IsNullOrWhiteSpace(found) && File.Exists(found);
                        }
                        catch { return false; }
                    }

                    string? EnsurePlaceholder(string dir, string fileOnly, Editor? log, out bool createdNew)
                    {
                        createdNew = false;
                        if (string.IsNullOrWhiteSpace(dir) || string.IsNullOrWhiteSpace(fileOnly))
                            return null;

                        var path = Path.Combine(dir, fileOnly);
                        var ext  = Path.GetExtension(fileOnly)?.ToLowerInvariant() ?? "";

                        if (File.Exists(path))
                            return path;

                        try
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

                            if (ext == ".png")
                            {
                                File.WriteAllBytes(path, OneByOnePng());
                                createdNew = true;
                            }
                            else if (ext == ".jpg" || ext == ".jpeg")
                            {
                                File.WriteAllBytes(path, OneByOneJpeg());
                                createdNew = true;
                            }
                            else if (ext == ".bmp")
                            {
                                File.WriteAllBytes(path, OneByOneBmp());
                                createdNew = true;
                            }
                            else
                            {
                                using (File.Create(path)) { }
                                createdNew = true;
                            }

                            createdPlaceholders.Add(path);
                            log?.WriteMessage($"\n[placeholder] created: {path}");
                            return path;
                        }
                        catch (System.Exception ex)
                        {
                            log?.WriteMessage($"\n[placeholder error] {path}: {ex.Message}");
                            return null;
                        }
                    }

                    // 1) Raster images
                    try
                    {
                        ObjectId imgDictId = RasterImageDef.GetImageDictionary(db);
                        if (!imgDictId.IsNull)
                        {
                            var imgDict = (DBDictionary)tr.GetObject(imgDictId, OpenMode.ForRead);
                            ed?.WriteMessage($"\n[INFO] Image dictionary entries: {imgDict.Count}");
                            foreach (DBDictionaryEntry kv in imgDict)
                            {
                                try
                                {
                                    var def = tr.GetObject(kv.Value, OpenMode.ForRead) as RasterImageDef;
                                    if (def == null) continue;

                                    imgSeen++;
                                    string old = def.SourceFileName;
                                    string fileOnly = FileOnly(old);
                                    ed?.WriteMessage($"\n[IMG] {kv.Key} : '{old}'");

                                    if (string.IsNullOrEmpty(old) || Same(old, fileOnly))
                                        continue;

                                    if (!CanResolve(fileOnly))
                                    {
                                        EnsurePlaceholder(dwgDir, fileOnly, ed, out _);
                                    }

                                    try
                                    {
                                        def.UpgradeOpen();
                                        try { if (def.IsLoaded) def.Unload(true); } catch { }
                                        def.SourceFileName = fileOnly;
                                        imgCount++;
                                        ed?.WriteMessage($"\nIMAGE: '{old}' -> '{fileOnly}'");
                                    }
                                    catch (System.Exception ex)
                                    {
                                        imgSkipped++;
                                        ed?.WriteMessage($"\n[IMAGE skip] '{old}' -> '{fileOnly}' : {ex.Message}");
                                    }
                                }
                                catch (System.Exception ex) { ed?.WriteMessage($"\n[IMAGE error] {kv.Key}: {ex.Message}"); }
                            }
                        }
                        else
                        {
                            ed?.WriteMessage("\n[INFO] No image dictionary.");
                        }
                    }
                    catch (System.Exception ex) { ed?.WriteMessage($"\n[Images section error] {ex.Message}"); }

                    // 2) Underlays
                    try
                    {
                        var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

                        void ProcessUnderlayDict(string dictName, ref int updated, ref int skipped, ref int seen)
                        {
                            try
                            {
                                if (!nod.Contains(dictName))
                                {
                                    ed?.WriteMessage($"\n[INFO] Underlay dict not found: {dictName}");
                                    return;
                                }
                                var subId = nod.GetAt(dictName);
                                var dict = (DBDictionary)tr.GetObject(subId, OpenMode.ForRead);
                                ed?.WriteMessage($"\n[INFO] {dictName} entries: {dict.Count}");
                                foreach (DBDictionaryEntry kv in dict)
                                {
                                    try
                                    {
                                        var def = tr.GetObject(kv.Value, OpenMode.ForRead) as UnderlayDefinition;
                                        if (def == null) continue;

                                        seen++;
                                        string old = def.SourceFileName;
                                        string fileOnly = FileOnly(old);
                                        ed?.WriteMessage($"\n[UNDERLAY] {dictName}/{kv.Key} : '{old}'");

                                        if (string.IsNullOrEmpty(old) || Same(old, fileOnly))
                                            continue;

                                        if (!CanResolve(fileOnly))
                                        {
                                            EnsurePlaceholder(dwgDir, fileOnly, ed, out _);
                                        }

                                        try
                                        {
                                            def.UpgradeOpen();
                                            try { def.Unload(); } catch { }
                                            def.SourceFileName = fileOnly;
                                            updated++;
                                            ed?.WriteMessage($"\nUNDERLAY[{dictName}]: '{old}' -> '{fileOnly}'");
                                        }
                                        catch (System.Exception ex)
                                        {
                                            skipped++;
                                            ed?.WriteMessage($"\n[UNDERLAY skip] {dictName}: '{old}' -> '{fileOnly}' : {ex.Message}");
                                        }
                                    }
                                    catch (System.Exception ex) { ed?.WriteMessage($"\n[UNDERLAY error] {dictName}/{kv.Key}: {ex.Message}"); }
                                }
                            }
                            catch (System.Exception ex) { ed?.WriteMessage($"\n[Underlay dict error] {dictName}: {ex.Message}"); }
                        }

                        ProcessUnderlayDict("ACAD_PDFDEFINITIONS", ref pdfCount, ref pdfSkipped, ref pdfSeen);
                        ProcessUnderlayDict("ACAD_DGNDEFINITIONS", ref dgnCount, ref dgnSkipped, ref dgnSeen);
                        ProcessUnderlayDict("ACAD_DWFDEFINITIONS", ref dwfCount, ref dwfSkipped, ref dwfSeen);
                    }
                    catch (System.Exception ex) { ed?.WriteMessage($"\n[Underlays section error] {ex.Message}"); }

                    // 3) DWG Xrefs
                    try
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        int totalBtrs = 0;
                        foreach (ObjectId btrId in bt) totalBtrs++;
                        ed?.WriteMessage($"\n[INFO] BlockTable records: {totalBtrs}");

                        foreach (ObjectId btrId in bt)
                        {
                            try
                            {
                                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                                if (btr == null) continue;

                                if (btr.IsFromExternalReference || btr.IsFromOverlayReference)
                                {
                                    xrefSeen++;
                                    string old = btr.PathName; // may be empty
                                    ed?.WriteMessage($"\n[XREF] {btr.Name} : '{old}'");

                                    if (!string.IsNullOrWhiteSpace(old))
                                    {
                                        string fn = FileOnly(old);
                                        if (!Same(old, fn))
                                        {
                                            btr.UpgradeOpen();
                                            btr.PathName = fn;
                                            xrefUpdated++;
                                            ed?.WriteMessage($"\nXREF: '{old}' -> '{fn}'");
                                        }
                                    }
                                }
                            }
                            catch (System.Exception ex) { ed?.WriteMessage($"\n[XREF error] BTR {btrId}: {ex.Message}"); }
                        }
                    }
                    catch (System.Exception ex) { ed?.WriteMessage($"\n[Xrefs section error] {ex.Message}"); }

                    tr.Commit();
                }

                if (DeletePlaceholdersAfterCommit)
                {
                    foreach (var p in createdPlaceholders)
                    {
                        try { if (File.Exists(p)) File.Delete(p); } catch { }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\n[FATAL] {ex}");
            }

            ed?.WriteMessage(
                $"\nDone. Seen: Images={imgSeen}, PDFs={pdfSeen}, DGN={dgnSeen}, DWF={dwfSeen}, Xrefs={xrefSeen}." +
                $"\nUpdated: Images={imgCount} (skipped {imgSkipped}), PDFs={pdfCount} (skipped {pdfSkipped}), " +
                $"DGN={dgnCount} (skipped {dgnSkipped}), DWF={dwfCount} (skipped {dwfSkipped}), Xrefs updated={xrefUpdated}.");
        }

        // helpers
        private static string FileOnly(string p)
        {
            if (string.IsNullOrWhiteSpace(p)) return p;
            try { return Path.GetFileName(p.Replace('/', '\\')); } catch { return p; }
        }
        private static bool Same(string a, string b) =>
            string.Equals(a?.Trim(), b?.Trim(), StringComparison.OrdinalIgnoreCase);

        private static string SafeDirName(string? path)
        {
            try { return Path.GetDirectoryName(path!) ?? Environment.CurrentDirectory; }
            catch { return Environment.CurrentDirectory; }
        }

        // tiny valid files
        private static byte[] OneByOnePng() => new byte[] {
            137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,137,0,0,0,10,73,68,65,84,120,156,99,248,15,4,0,9,251,3,253,160,119,103,145,0,0,0,0,73,69,78,68,174,66,96,130
        };
        private static byte[] OneByOneJpeg() => Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBxAQEBUQFRUVFRUVFRUVFRUVFRUVFRUWFhUVFRUYHSggGBolGxUVITEhJSkrLi4uFx8zODMtNygtLisBCgoKDg0OGhAQGy0lICUtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLf/AABEIAAEAAQMBIgACEQEDEQH/xAAbAAEAAQUBAAAAAAAAAAAAAAADBAECBgf/xAAcEAACAgMBAQAAAAAAAAAAAAAAAQIDBBEFEhP/xAAaAQACAgMAAAAAAAAAAAAAAAAAAgEDBBEF/8QAGhEAAQUBAAAAAAAAAAAAAAAAAQACAxITIf/aAAwDAQACEQMRAD8A7oAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/Z"
        );
        private static byte[] OneByOneBmp()
        {
            var bytes = new byte[58];
            bytes[0] = (byte)'B'; bytes[1] = (byte)'M';
            int fileSize = 58; Bit(bytes, 2, fileSize);
            Bit(bytes, 10, 54);
            Bit(bytes, 14, 40);
            Bit(bytes, 18, 1);
            Bit(bytes, 22, 1);
            bytes[26] = 1; bytes[27] = 0;
            bytes[28] = 24; bytes[29] = 0;
            bytes[54] = 255; bytes[55] = 0; bytes[56] = 0; bytes[57] = 0;
            return bytes;

            static void Bit(byte[] b, int i, int v)
            {
                b[i+0]=(byte)(v & 0xFF);
                b[i+1]=(byte)((v>>8) & 0xFF);
                b[i+2]=(byte)((v>>16)& 0xFF);
                b[i+3]=(byte)((v>>24)& 0xFF);
            }
        }
    }
}
