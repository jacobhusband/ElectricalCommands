using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Acies.AutoCAD.Shared
{
    internal sealed class StoredTitleBlockPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

        internal Point3d ToPoint3d()
        {
            return new Point3d(X, Y, Z);
        }

        internal static StoredTitleBlockPoint FromPoint3d(Point3d point)
        {
            return new StoredTitleBlockPoint
            {
                X = point.X,
                Y = point.Y,
                Z = point.Z
            };
        }
    }

    internal sealed class ProjectTitleBlockSettings
    {
        public int Version { get; set; } = 1;
        public StoredTitleBlockPoint LowerLeft { get; set; }
        public StoredTitleBlockPoint UpperRight { get; set; }
        public string SourceDrawing { get; set; } = string.Empty;
        public string SourceSpace { get; set; } = string.Empty;
        public string UpdatedUtc { get; set; } = string.Empty;

        internal double Width => UpperRight == null || LowerLeft == null
            ? 0.0
            : Math.Abs(UpperRight.X - LowerLeft.X);

        internal double Height => UpperRight == null || LowerLeft == null
            ? 0.0
            : Math.Abs(UpperRight.Y - LowerLeft.Y);

        internal bool IsValid =>
            LowerLeft != null &&
            UpperRight != null &&
            Width > 1e-6 &&
            Height > 1e-6;

        internal Point3d[] ToBoundaryPolygon()
        {
            if (!IsValid) return null;

            double minX = Math.Min(LowerLeft.X, UpperRight.X);
            double minY = Math.Min(LowerLeft.Y, UpperRight.Y);
            double maxX = Math.Max(LowerLeft.X, UpperRight.X);
            double maxY = Math.Max(LowerLeft.Y, UpperRight.Y);
            const double z = 0.0;

            return new[]
            {
                new Point3d(minX, minY, z),
                new Point3d(maxX, minY, z),
                new Point3d(maxX, maxY, z),
                new Point3d(minX, maxY, z)
            };
        }
    }

    internal sealed class SheetSizeProfile
    {
        internal SheetSizeProfile(string key, string displayName, double shortSideInches, double longSideInches)
        {
            Key = key;
            DisplayName = displayName;
            ShortSideInches = shortSideInches;
            LongSideInches = longSideInches;
        }

        internal string Key { get; }
        internal string DisplayName { get; }
        internal double ShortSideInches { get; }
        internal double LongSideInches { get; }
    }

    internal static class ProjectTitleBlockSettingsStore
    {
        internal const string SetCommandName = "SETTITLEBLOCK";

        private const string SettingsDirectoryName = ".acies";
        private const string SettingsFileName = "titleblock-boundary.json";

        private static readonly string[] ProjectTopLevelFolders =
        {
            "Arch",
            "Archive",
            "BGD",
            "CAD Release",
            "Documents",
            "Electrical",
            "Mechanical",
            "PDF",
            "Plumbing",
            "RFI",
            "Submittals",
            "Survey",
            "Xrefs"
        };

        // These are the two folders used by this workflow: sheets normally live under
        // Electrical and the titleblock source normally lives under Xrefs. Generic
        // names such as "Documents" are deliberately not treated as project markers.
        private static readonly string[] DistinctiveProjectTopLevelFolders =
        {
            "Electrical",
            "Xrefs"
        };

        private static readonly SheetSizeProfile[] Profiles =
        {
            new SheetSizeProfile("ANSI_D", "22 x 34 (ANSI D)", 22.0, 34.0),
            new SheetSizeProfile("ARCH_D", "24 x 36 (ARCH D)", 24.0, 36.0),
            new SheetSizeProfile("ARCH_E1", "30 x 42 (ARCH E1)", 30.0, 42.0),
            new SheetSizeProfile("ARCH_E", "36 x 48 (ARCH E)", 36.0, 48.0)
        };

        internal static bool TryLoad(
            Document doc,
            Database db,
            out ProjectTitleBlockSettings settings,
            out string projectRoot,
            out string settingsPath,
            out string failureReason)
        {
            settings = null;
            projectRoot = string.Empty;
            settingsPath = string.Empty;
            failureReason = string.Empty;

            if (!TryResolveProjectContext(doc, db, out projectRoot, out settingsPath, out failureReason))
            {
                return false;
            }

            if (!File.Exists(settingsPath))
            {
                failureReason = $"No saved project titleblock boundary was found at '{settingsPath}'.";
                return false;
            }

            try
            {
                string json = File.ReadAllText(settingsPath);
                settings = JsonConvert.DeserializeObject<ProjectTitleBlockSettings>(json);
                if (settings == null || !settings.IsValid)
                {
                    settings = null;
                    failureReason = $"The saved titleblock boundary is invalid: '{settingsPath}'.";
                    return false;
                }

                return true;
            }
            catch (System.Exception ex)
            {
                settings = null;
                failureReason = $"Unable to read the saved titleblock boundary: {ex.Message}";
                return false;
            }
        }

        internal static bool TryPromptAndSave(
            Document doc,
            Editor ed,
            Database db,
            out ProjectTitleBlockSettings settings,
            out string projectRoot,
            out string settingsPath)
        {
            settings = null;
            projectRoot = string.Empty;
            settingsPath = string.Empty;

            if (doc == null || ed == null || db == null)
            {
                return false;
            }

            if (!TryResolveProjectContext(doc, db, out projectRoot, out settingsPath, out string contextFailure))
            {
                ed.WriteMessage($"\n{SetCommandName}: {contextFailure}");
                return false;
            }

            var firstOptions = new PromptPointOptions("\nSelect the first corner of the titleblock boundary: ");
            PromptPointResult firstResult = ed.GetPoint(firstOptions);
            if (firstResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage($"\n{SetCommandName} cancelled.");
                return false;
            }

            var oppositeOptions = new PromptCornerOptions(
                "\nSelect the opposite corner of the titleblock boundary: ",
                firstResult.Value)
            {
                UseDashedLine = true
            };

            PromptPointResult oppositeResult = ed.GetCorner(oppositeOptions);
            if (oppositeResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage($"\n{SetCommandName} cancelled.");
                return false;
            }

            double minX = Math.Min(firstResult.Value.X, oppositeResult.Value.X);
            double minY = Math.Min(firstResult.Value.Y, oppositeResult.Value.Y);
            double maxX = Math.Max(firstResult.Value.X, oppositeResult.Value.X);
            double maxY = Math.Max(firstResult.Value.Y, oppositeResult.Value.Y);

            if (Math.Abs(maxX - minX) <= 1e-6 || Math.Abs(maxY - minY) <= 1e-6)
            {
                ed.WriteMessage("\nThe selected titleblock boundary must have a non-zero width and height.");
                return false;
            }

            string sourceSpace = string.Empty;
            try
            {
                sourceSpace = db.TileMode
                    ? "Model"
                    : (LayoutManager.Current?.CurrentLayout ?? "Paper");
            }
            catch
            {
                sourceSpace = db.TileMode ? "Model" : "Paper";
            }

            settings = new ProjectTitleBlockSettings
            {
                Version = 1,
                LowerLeft = StoredTitleBlockPoint.FromPoint3d(new Point3d(minX, minY, firstResult.Value.Z)),
                UpperRight = StoredTitleBlockPoint.FromPoint3d(new Point3d(maxX, maxY, oppositeResult.Value.Z)),
                SourceDrawing = ResolveBestDrawingPath(doc, db),
                SourceSpace = sourceSpace,
                UpdatedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
            };

            if (!TrySave(settingsPath, settings, out string saveFailure))
            {
                settings = null;
                ed.WriteMessage($"\n{SetCommandName}: {saveFailure}");
                return false;
            }

            return true;
        }

        internal static bool TrySaveBoundary(
            Document doc,
            Database db,
            Point3d[] boundary,
            out string projectRoot,
            out string settingsPath,
            out string failureReason)
        {
            projectRoot = string.Empty;
            settingsPath = string.Empty;
            failureReason = string.Empty;

            if (boundary == null || boundary.Length < 2)
            {
                failureReason = "The titleblock boundary is empty.";
                return false;
            }

            if (!TryResolveProjectContext(doc, db, out projectRoot, out settingsPath, out failureReason))
            {
                return false;
            }

            double minX = boundary.Min(point => point.X);
            double minY = boundary.Min(point => point.Y);
            double maxX = boundary.Max(point => point.X);
            double maxY = boundary.Max(point => point.Y);
            if (Math.Abs(maxX - minX) <= 1e-6 || Math.Abs(maxY - minY) <= 1e-6)
            {
                failureReason = "The titleblock boundary must have a non-zero width and height.";
                return false;
            }

            string sourceSpace = string.Empty;
            try
            {
                sourceSpace = db.TileMode
                    ? "Model"
                    : (LayoutManager.Current?.CurrentLayout ?? "Paper");
            }
            catch
            {
                sourceSpace = db.TileMode ? "Model" : "Paper";
            }

            var settings = new ProjectTitleBlockSettings
            {
                Version = 1,
                LowerLeft = StoredTitleBlockPoint.FromPoint3d(new Point3d(minX, minY, boundary[0].Z)),
                UpperRight = StoredTitleBlockPoint.FromPoint3d(new Point3d(maxX, maxY, boundary[0].Z)),
                SourceDrawing = ResolveBestDrawingPath(doc, db),
                SourceSpace = sourceSpace,
                UpdatedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
            };

            return TrySave(settingsPath, settings, out failureReason);
        }

        internal static bool TryClassifySheetSize(
            ProjectTitleBlockSettings settings,
            out SheetSizeProfile profile,
            out string failureReason)
        {
            profile = null;
            failureReason = string.Empty;
            if (settings == null || !settings.IsValid)
            {
                failureReason = "The saved titleblock boundary is invalid.";
                return false;
            }

            double observedShort = Math.Min(settings.Width, settings.Height);
            double observedLong = Math.Max(settings.Width, settings.Height);

            var scaleCandidates = new[]
            {
                new { Name = "inches", ToInches = 1.0 },
                new { Name = "millimeters", ToInches = 1.0 / 25.4 }
            };

            var matches = new List<Tuple<SheetSizeProfile, double, double, string>>();
            foreach (var scale in scaleCandidates)
            {
                double shortInches = observedShort * scale.ToInches;
                double longInches = observedLong * scale.ToInches;

                foreach (SheetSizeProfile candidate in Profiles)
                {
                    double shortError = Math.Abs(shortInches - candidate.ShortSideInches);
                    double longError = Math.Abs(longInches - candidate.LongSideInches);
                    matches.Add(Tuple.Create(candidate, shortError + longError, Math.Max(shortError, longError), scale.Name));
                }
            }

            var sorted = matches
                .OrderBy(match => match.Item2)
                .ThenBy(match => match.Item3)
                .ToList();

            var best = sorted[0];
            if (best.Item3 > 1.5)
            {
                failureReason =
                    $"Saved boundary {settings.Width:0.###} x {settings.Height:0.###} does not match a supported sheet size " +
                    "(22x34, 24x36, 30x42, or 36x48). Run SETTITLEBLOCK again and select the outer sheet boundary.";
                return false;
            }

            var runnerUp = sorted.Skip(1).FirstOrDefault(match => !ReferenceEquals(match.Item1, best.Item1));
            if (runnerUp != null && runnerUp.Item2 - best.Item2 < 1.0)
            {
                failureReason =
                    $"Saved boundary {settings.Width:0.###} x {settings.Height:0.###} is ambiguous. " +
                    "Run SETTITLEBLOCK again and select the outer sheet boundary.";
                return false;
            }

            profile = best.Item1;
            return true;
        }

        internal static bool TryResolveProjectContext(
            Document doc,
            Database db,
            out string projectRoot,
            out string settingsPath,
            out string failureReason)
        {
            projectRoot = string.Empty;
            settingsPath = string.Empty;
            failureReason = string.Empty;

            string drawingPath = ResolveBestDrawingPath(doc, db);
            if (string.IsNullOrWhiteSpace(drawingPath))
            {
                failureReason = "Save the active drawing inside the project folder first.";
                return false;
            }

            string drawingFolder;
            try
            {
                drawingFolder = Path.GetDirectoryName(drawingPath);
            }
            catch (System.Exception ex)
            {
                failureReason = $"Unable to resolve the drawing folder: {ex.Message}";
                return false;
            }

            if (string.IsNullOrWhiteSpace(drawingFolder) || !Directory.Exists(drawingFolder))
            {
                failureReason = "The active drawing folder does not exist.";
                return false;
            }

            projectRoot = ResolveProjectRoot(drawingFolder);
            if (string.IsNullOrWhiteSpace(projectRoot))
            {
                failureReason = "Unable to determine the project root folder.";
                return false;
            }

            settingsPath = Path.Combine(projectRoot, SettingsDirectoryName, SettingsFileName);
            return true;
        }

        private static string ResolveBestDrawingPath(Document doc, Database db)
        {
            var candidates = new List<string>();
            if (doc != null && !string.IsNullOrWhiteSpace(doc.Name))
            {
                candidates.Add(doc.Name);
            }
            if (db != null && !string.IsNullOrWhiteSpace(db.Filename))
            {
                candidates.Add(db.Filename);
            }

            foreach (string candidate in candidates)
            {
                try
                {
                    string normalized = Path.GetFullPath(candidate.Trim().Trim('"'));
                    if (!Path.IsPathRooted(normalized)) continue;
                    if (!string.Equals(Path.GetExtension(normalized), ".dwg", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!Directory.Exists(Path.GetDirectoryName(normalized))) continue;
                    if (!File.Exists(normalized)) continue;
                    return normalized;
                }
                catch
                {
                    // Try the next candidate.
                }
            }

            return string.Empty;
        }

        private static string ResolveProjectRoot(string drawingFolder)
        {
            DirectoryInfo current;
            try
            {
                current = new DirectoryInfo(drawingFolder);
            }
            catch
            {
                return drawingFolder;
            }

            DirectoryInfo fallback = current;
            for (DirectoryInfo cursor = current; cursor != null; cursor = cursor.Parent)
            {
                if (DistinctiveProjectTopLevelFolders.Any(folder =>
                    string.Equals(folder, cursor.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    return cursor.Parent?.FullName ?? cursor.FullName;
                }

                int markerCount = 0;
                foreach (string folder in ProjectTopLevelFolders)
                {
                    try
                    {
                        if (Directory.Exists(Path.Combine(cursor.FullName, folder)))
                        {
                            markerCount++;
                        }
                    }
                    catch
                    {
                        // Ignore an inaccessible marker and continue scoring this ancestor.
                    }
                }

                if (markerCount >= 2)
                {
                    return cursor.FullName;
                }
            }

            return fallback.FullName;
        }

        private static bool TrySave(
            string settingsPath,
            ProjectTitleBlockSettings settings,
            out string failureReason)
        {
            failureReason = string.Empty;
            try
            {
                string settingsDirectory = Path.GetDirectoryName(settingsPath);
                if (string.IsNullOrWhiteSpace(settingsDirectory))
                {
                    failureReason = "Unable to determine the settings directory.";
                    return false;
                }

                Directory.CreateDirectory(settingsDirectory);
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                string temporaryPath = settingsPath + "." + Guid.NewGuid().ToString("N") + ".tmp";

                try
                {
                    File.WriteAllText(temporaryPath, json);
                    if (File.Exists(settingsPath))
                    {
                        try
                        {
                            File.Replace(temporaryPath, settingsPath, null);
                        }
                        catch (IOException)
                        {
                            // Some project network shares do not support atomic replace.
                            File.Copy(temporaryPath, settingsPath, true);
                            File.Delete(temporaryPath);
                        }
                        catch (PlatformNotSupportedException)
                        {
                            File.Copy(temporaryPath, settingsPath, true);
                            File.Delete(temporaryPath);
                        }
                    }
                    else
                    {
                        File.Move(temporaryPath, settingsPath);
                    }
                }
                finally
                {
                    if (File.Exists(temporaryPath))
                    {
                        try { File.Delete(temporaryPath); } catch { }
                    }
                }

                return true;
            }
            catch (System.Exception ex)
            {
                failureReason = $"Unable to save the project titleblock boundary: {ex.Message}";
                return false;
            }
        }
    }
}
