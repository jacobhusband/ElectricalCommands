using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoCADCleanupTool
{
    internal enum TitleBlockResolutionKind
    {
        Resolved = 0,
        Ambiguous = 1,
        NotFound = 2
    }

    internal sealed class TitleBlockXrefCandidate
    {
        internal ObjectId XrefBtrId { get; set; } = ObjectId.Null;
        internal ObjectId BlockReferenceId { get; set; } = ObjectId.Null;
        internal string LayoutName { get; set; } = string.Empty;
        internal string BlockName { get; set; } = string.Empty;
        internal string PathName { get; set; } = string.Empty;
        internal Point3d InsertionPoint { get; set; } = Point3d.Origin;
        internal int LayoutXrefCount { get; set; }
        internal int Score { get; set; }
    }

    internal sealed class TitleBlockXrefResolutionResult
    {
        internal TitleBlockXrefResolutionResult(
            TitleBlockResolutionKind kind,
            TitleBlockXrefCandidate winner,
            IReadOnlyList<TitleBlockXrefCandidate> candidates)
        {
            Kind = kind;
            Winner = winner;
            Candidates = candidates ?? Array.Empty<TitleBlockXrefCandidate>();
        }

        internal TitleBlockResolutionKind Kind { get; }
        internal TitleBlockXrefCandidate Winner { get; }
        internal IReadOnlyList<TitleBlockXrefCandidate> Candidates { get; }
    }

    internal static class TitleBlockXrefResolver
    {
        private static readonly string[] _hints = { "x-tb", "title", "tblock", "border", "sheet" };

        internal static IReadOnlyList<TitleBlockXrefCandidate> GetLikelyTitleBlockCandidates(Database db)
        {
            List<TitleBlockXrefCandidate> sorted = GetScoredCandidates(db);
            if (sorted.Count == 0) return Array.Empty<TitleBlockXrefCandidate>();

            int best = sorted[0].Score;
            int threshold = Math.Max(80, best - 25);

            return sorted
                .Where(c => c.Score >= threshold)
                .ToList();
        }

        internal static TitleBlockXrefResolutionResult Resolve(Database db)
        {
            if (db == null)
            {
                return new TitleBlockXrefResolutionResult(TitleBlockResolutionKind.NotFound, null, Array.Empty<TitleBlockXrefCandidate>());
            }

            List<TitleBlockXrefCandidate> sorted = GetScoredCandidates(db);
            if (sorted.Count == 0)
            {
                return new TitleBlockXrefResolutionResult(TitleBlockResolutionKind.NotFound, null, sorted);
            }

            TitleBlockXrefCandidate best = sorted[0];
            int runnerUpScore = sorted.Count > 1 ? sorted[1].Score : int.MinValue;
            int scoreDelta = sorted.Count > 1 ? best.Score - runnerUpScore : best.Score;

            if (best.Score < 40)
            {
                return new TitleBlockXrefResolutionResult(TitleBlockResolutionKind.NotFound, null, sorted);
            }

            if (sorted.Count == 1)
            {
                return new TitleBlockXrefResolutionResult(TitleBlockResolutionKind.Resolved, best, sorted);
            }

            bool confident =
                (best.Score >= 95 && scoreDelta >= 25) ||
                (best.Score >= 80 && best.LayoutXrefCount == 1 && scoreDelta >= 15) ||
                (best.Score >= 150 && scoreDelta >= 10);

            if (confident)
            {
                return new TitleBlockXrefResolutionResult(TitleBlockResolutionKind.Resolved, best, sorted);
            }

            return new TitleBlockXrefResolutionResult(TitleBlockResolutionKind.Ambiguous, null, sorted);
        }

        private static List<TitleBlockXrefCandidate> GetScoredCandidates(Database db)
        {
            List<TitleBlockXrefCandidate> candidates = CollectPaperSpaceXrefCandidates(db);
            if (candidates.Count == 0) return candidates;

            string currentLayout = string.Empty;
            try
            {
                currentLayout = LayoutManager.Current?.CurrentLayout ?? string.Empty;
            }
            catch { }

            var layoutCounts = candidates
                .GroupBy(c => c.LayoutName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);

            foreach (var candidate in candidates)
            {
                candidate.LayoutXrefCount = layoutCounts.TryGetValue(candidate.LayoutName ?? string.Empty, out int count) ? count : 0;
                candidate.Score = ScoreCandidate(candidate, currentLayout);
            }

            return candidates
                .OrderByDescending(c => c.Score)
                .ThenBy(c => c.LayoutName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(c => c.BlockName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<TitleBlockXrefCandidate> CollectPaperSpaceXrefCandidates(Database db)
        {
            var candidates = new List<TitleBlockXrefCandidate>();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (layoutDict == null)
                    {
                        tr.Commit();
                        return candidates;
                    }

                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                        if (layout == null || layout.ModelType) continue;

                        var layoutBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                        if (layoutBtr == null) continue;

                        foreach (ObjectId entId in layoutBtr)
                        {
                            var br = tr.GetObject(entId, OpenMode.ForRead, false) as BlockReference;
                            if (br == null) continue;

                            var def = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead, false) as BlockTableRecord;
                            if (def == null) continue;
                            if (!def.IsFromExternalReference && !def.IsFromOverlayReference) continue;

                            candidates.Add(new TitleBlockXrefCandidate
                            {
                                XrefBtrId = br.BlockTableRecord,
                                BlockReferenceId = br.ObjectId,
                                LayoutName = layout.LayoutName ?? string.Empty,
                                BlockName = def.Name ?? string.Empty,
                                PathName = def.PathName ?? string.Empty,
                                InsertionPoint = br.Position
                            });
                        }
                    }

                    tr.Commit();
                }
            }
            catch
            {
                // Return what we collected so far; caller can choose fallback behavior.
            }

            return candidates;
        }

        private static int ScoreCandidate(TitleBlockXrefCandidate candidate, string currentLayout)
        {
            int score = 0;
            string name = (candidate.BlockName ?? string.Empty).ToLowerInvariant();
            string path = (candidate.PathName ?? string.Empty).ToLowerInvariant();
            string normalizedPath = path.Replace('/', '\\');
            string fileNoExt = string.Empty;
            string blockToken = CleanupCommands.CanonicalizeXrefToken(candidate.BlockName);
            string fileToken = string.Empty;
            try { fileToken = CleanupCommands.CanonicalizeXrefToken(Path.GetFileName(candidate.PathName ?? string.Empty)); } catch { }
            bool exactXtb = IsExactXtbToken(blockToken) || IsExactXtbToken(fileToken);

            try { fileNoExt = Path.GetFileNameWithoutExtension(path)?.ToLowerInvariant() ?? string.Empty; } catch { }

            if (!string.IsNullOrWhiteSpace(currentLayout) &&
                string.Equals(candidate.LayoutName, currentLayout, StringComparison.OrdinalIgnoreCase))
            {
                score += 40;
            }

            if (ContainsHint(name) || ContainsHint(fileNoExt))
            {
                score += 130;
            }

            if (ContainsHint(path))
            {
                score += 120;
            }

            if (exactXtb)
            {
                score += 220;
            }

            bool inXrefsFolder = normalizedPath.Contains("\\xrefs\\") ||
                                 normalizedPath.StartsWith("xrefs\\", StringComparison.OrdinalIgnoreCase) ||
                                 normalizedPath.StartsWith("..\\xrefs\\", StringComparison.OrdinalIgnoreCase);
            bool hasXtbInPath = normalizedPath.Contains("x-tb") || normalizedPath.Contains("x_tb");
            if (inXrefsFolder && (exactXtb || hasXtbInPath))
            {
                score += 140;
            }

            if (candidate.LayoutXrefCount == 1)
            {
                score += 70;
            }
            else if (candidate.LayoutXrefCount == 2)
            {
                score += 20;
            }

            double distanceToOrigin = Math.Sqrt(
                (candidate.InsertionPoint.X * candidate.InsertionPoint.X) +
                (candidate.InsertionPoint.Y * candidate.InsertionPoint.Y));

            if (distanceToOrigin <= 1e-4)
            {
                score += 25;
            }
            else if (distanceToOrigin <= 1.0)
            {
                score += 16;
            }
            else if (distanceToOrigin <= 24.0)
            {
                score += 8;
            }

            if (name.StartsWith("x-", StringComparison.OrdinalIgnoreCase))
            {
                score += 8;
            }

            if (IsLikelyDwg(path, name))
            {
                score += 10;
            }

            return score;
        }

        private static bool IsExactXtbToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            string normalized = token.Trim().ToLowerInvariant();
            return string.Equals(normalized, "x-tb", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "x_tb", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "xtb", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsHint(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return false;
            foreach (string hint in _hints)
            {
                if (input.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsLikelyDwg(string path, string blockName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) &&
                    string.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch { }

            return (blockName ?? string.Empty).EndsWith(".dwg", StringComparison.OrdinalIgnoreCase);
        }
    }
}
