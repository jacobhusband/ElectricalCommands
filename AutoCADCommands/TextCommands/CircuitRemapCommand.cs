using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  // Moves a run of circuits to a new panel/number in one pass, e.g. DP-16..36 (evens) -> DP2-2..22.
  // Every TEXT, MTEXT, attribute, and multileader label is rewritten simultaneously, so a remap onto
  // the same panel (DP-16 -> DP-2) never chains into a second rewrite.
  public static class CircuitRemapCommand
  {
    private static string _lastOldPanel = string.Empty;
    private static int _lastOldFirst = 1;
    private static int _lastOldLast = 1;
    private static string _lastNewPanel = string.Empty;
    private static int _lastNewFirst = 1;
    private static bool _lastSameSide = true;
    private static bool _lastWholeDrawing = true;

    [CommandMethod("CIRCUITREMAP", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("CREMAP", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void CircuitRemap()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        return;
      }

      try
      {
        PromptSelectionResult implied = ed.SelectImplied();

        string oldPanel = PromptPanel(ed, "\nOld panel name", _lastOldPanel);
        if (oldPanel == null) return;
        int? oldFirst = PromptCircuit(ed, "\nFirst old circuit", _lastOldFirst);
        if (oldFirst == null) return;
        int? oldLast = PromptCircuit(ed, "\nLast old circuit", Math.Max(_lastOldLast, oldFirst.Value));
        if (oldLast == null) return;
        if (oldLast.Value < oldFirst.Value)
        {
          ed.WriteMessage("\nThe last circuit must not be lower than the first circuit.");
          return;
        }

        bool? sameSide = PromptSameSide(ed, _lastSameSide);
        if (sameSide == null) return;
        int step = sameSide.Value ? 2 : 1;
        if (step == 2 && (oldLast.Value - oldFirst.Value) % 2 != 0)
        {
          ed.WriteMessage($"\nCircuits {oldFirst.Value} and {oldLast.Value} are not on the same side of the panel.");
          return;
        }

        string newPanel = PromptPanel(ed, "\nNew panel name", string.IsNullOrEmpty(_lastNewPanel) ? oldPanel : _lastNewPanel);
        if (newPanel == null) return;
        int? newFirst = PromptCircuit(ed, "\nFirst new circuit", _lastNewFirst);
        if (newFirst == null) return;

        CircuitMap map = new CircuitMap(oldPanel, oldFirst.Value, oldLast.Value, step, newPanel, newFirst.Value);
        if (map.Count == 0)
        {
          ed.WriteMessage("\nNo circuits fall in that range.");
          return;
        }
        if (step == 2 && newFirst.Value % 2 != oldFirst.Value % 2)
        {
          ed.WriteMessage("\nNote: the new circuits are on the opposite side of the panel from the old ones.");
        }

        _lastOldPanel = oldPanel;
        _lastOldFirst = oldFirst.Value;
        _lastOldLast = oldLast.Value;
        _lastSameSide = sameSide.Value;
        _lastNewPanel = newPanel;
        _lastNewFirst = newFirst.Value;

        ed.WriteMessage("\nCircuit map:");
        foreach (KeyValuePair<int, int> pair in map.Pairs)
        {
          ed.WriteMessage($"\n  {map.OldPanel}-{pair.Key} -> {map.NewPanel}-{pair.Value}");
        }

        IEnumerable<ObjectId> scopeIds;
        if (implied.Status == PromptStatus.OK && implied.Value != null && implied.Value.Count > 0)
        {
          scopeIds = implied.Value.GetObjectIds();
        }
        else
        {
          bool? wholeDrawing = PromptScope(ed, _lastWholeDrawing);
          if (wholeDrawing == null) return;
          _lastWholeDrawing = wholeDrawing.Value;
          if (wholeDrawing.Value)
          {
            scopeIds = null;
          }
          else
          {
            PromptSelectionOptions options = new PromptSelectionOptions
            {
              MessageForAdding = "\nSelect circuit labels: ",
            };
            SelectionFilter filter = new SelectionFilter(new[]
            {
              new TypedValue((int)DxfCode.Start, "TEXT,MTEXT,INSERT,MULTILEADER"),
            });
            PromptSelectionResult selection = ed.GetSelection(options, filter);
            if (selection.Status != PromptStatus.OK || selection.Value == null)
            {
              return;
            }
            scopeIds = selection.Value.GetObjectIds();
          }
        }

        RemapPlan plan = BuildPlan(db, scopeIds, map);
        ReportPlan(ed, map, plan);
        if (plan.Edits.Count == 0)
        {
          return;
        }

        PromptKeywordOptions confirm = new PromptKeywordOptions($"\nUpdate {plan.Edits.Count} object(s)? [Yes/No]", "Yes No")
        {
          AllowNone = true,
        };
        confirm.Keywords.Default = "Yes";
        PromptResult confirmResult = ed.GetKeywords(confirm);
        if (confirmResult.Status != PromptStatus.OK || confirmResult.StringResult != "Yes")
        {
          ed.WriteMessage("\nCIRCUITREMAP canceled. Nothing was changed.");
          return;
        }

        ApplyPlan(db, plan);
        ed.WriteMessage($"\nCIRCUITREMAP updated {plan.Edits.Count} object(s). Update the panel schedule to match.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("\nCIRCUITREMAP error: " + ex.Message);
      }
      finally
      {
        ed.SetImpliedSelection(Array.Empty<ObjectId>());
      }
    }

    private static string PromptPanel(Editor ed, string message, string defaultValue)
    {
      PromptStringOptions options = new PromptStringOptions(
        string.IsNullOrEmpty(defaultValue) ? message + ": " : $"{message} <{defaultValue}>: ")
      {
        AllowSpaces = false,
      };
      PromptResult result = ed.GetString(options);
      if (result.Status != PromptStatus.OK)
      {
        return null;
      }

      string value = NormalizePanel(result.StringResult);
      if (value.Length == 0)
      {
        value = defaultValue ?? string.Empty;
      }
      if (value.Length == 0)
      {
        ed.WriteMessage("\nA panel name is required.");
        return null;
      }
      return value;
    }

    private static int? PromptCircuit(Editor ed, string message, int defaultValue)
    {
      PromptIntegerOptions options = new PromptIntegerOptions(message)
      {
        AllowNegative = false,
        AllowZero = false,
        AllowNone = true,
        DefaultValue = Math.Max(1, defaultValue),
        UseDefaultValue = true,
      };
      PromptIntegerResult result = ed.GetInteger(options);
      return result.Status == PromptStatus.OK ? result.Value : (int?)null;
    }

    private static bool? PromptSameSide(Editor ed, bool defaultSameSide)
    {
      PromptKeywordOptions options = new PromptKeywordOptions("\nCircuits to move [SameSide/All]", "SameSide All")
      {
        AllowNone = true,
      };
      options.Keywords.Default = defaultSameSide ? "SameSide" : "All";
      PromptResult result = ed.GetKeywords(options);
      return result.Status == PromptStatus.OK ? result.StringResult == "SameSide" : (bool?)null;
    }

    private static bool? PromptScope(Editor ed, bool defaultWholeDrawing)
    {
      PromptKeywordOptions options = new PromptKeywordOptions("\nApply to [Drawing/Select]", "Drawing Select")
      {
        AllowNone = true,
      };
      options.Keywords.Default = defaultWholeDrawing ? "Drawing" : "Select";
      PromptResult result = ed.GetKeywords(options);
      return result.Status == PromptStatus.OK ? result.StringResult == "Drawing" : (bool?)null;
    }

    internal static string NormalizePanel(string value)
    {
      return (value ?? string.Empty).Trim().TrimEnd('-').Trim().ToUpperInvariant();
    }

    private static RemapPlan BuildPlan(Database db, IEnumerable<ObjectId> scopeIds, CircuitMap map)
    {
      RemapPlan plan = new RemapPlan();
      using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
      {
        HashSet<ObjectId> lockedLayers = FindLockedLayers(tr, db);
        IEnumerable<ObjectId> ids = scopeIds ?? EnumerateLayoutEntities(tr, db);
        foreach (ObjectId id in ids)
        {
          Entity entity = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
          if (entity == null)
          {
            continue;
          }

          PlanEntity(tr, entity, map, lockedLayers, plan);
          if (entity is BlockReference blockReference)
          {
            foreach (ObjectId attributeId in blockReference.AttributeCollection)
            {
              AttributeReference attribute = tr.GetObject(attributeId, OpenMode.ForRead, false) as AttributeReference;
              if (attribute != null)
              {
                PlanEntity(tr, attribute, map, lockedLayers, plan);
              }
            }
          }
        }
        tr.Commit();
      }
      return plan;
    }

    private static void PlanEntity(
      Transaction tr,
      Entity entity,
      CircuitMap map,
      HashSet<ObjectId> lockedLayers,
      RemapPlan plan)
    {
      string original;
      TextKind kind;
      if (entity is AttributeReference attribute)
      {
        if (attribute.Invisible)
        {
          return;
        }
        kind = attribute.IsMTextAttribute ? TextKind.MTextAttribute : TextKind.Text;
        original = attribute.IsMTextAttribute ? attribute.MTextAttribute.Contents : attribute.TextString;
      }
      else if (entity is DBText text)
      {
        kind = TextKind.Text;
        original = text.TextString;
      }
      else if (entity is MText mtext)
      {
        kind = TextKind.MText;
        original = mtext.Contents;
      }
      else if (entity is MLeader leader && leader.ContentType == ContentType.MTextContent && leader.MText != null)
      {
        kind = TextKind.Leader;
        original = leader.MText.Contents;
      }
      else
      {
        return;
      }

      Dictionary<int, int> counts = new Dictionary<int, int>();
      string updated = map.Apply(original, counts, out List<string> partials);
      plan.Partials.AddRange(partials);
      if (updated == original)
      {
        return;
      }

      if (lockedLayers.Contains(entity.LayerId))
      {
        plan.LockedCount++;
        return;
      }

      plan.Edits.Add(new RemapEdit { Id = entity.ObjectId, Kind = kind, Contents = updated });
      foreach (KeyValuePair<int, int> pair in counts)
      {
        plan.Counts.TryGetValue(pair.Key, out int count);
        plan.Counts[pair.Key] = count + pair.Value;
      }
    }

    private static void ApplyPlan(Database db, RemapPlan plan)
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        foreach (RemapEdit edit in plan.Edits)
        {
          Entity entity = (Entity)tr.GetObject(edit.Id, OpenMode.ForWrite, false);
          switch (edit.Kind)
          {
            case TextKind.Text:
              ((DBText)entity).TextString = edit.Contents;
              break;
            case TextKind.MText:
              ((MText)entity).Contents = edit.Contents;
              break;
            case TextKind.MTextAttribute:
              AttributeReference attribute = (AttributeReference)entity;
              MText attributeText = attribute.MTextAttribute;
              attributeText.Contents = edit.Contents;
              attribute.MTextAttribute = attributeText;
              attribute.UpdateMTextAttribute();
              break;
            case TextKind.Leader:
              MLeader leader = (MLeader)entity;
              MText leaderText = leader.MText;
              leaderText.Contents = edit.Contents;
              leader.MText = leaderText;
              break;
          }
        }
        tr.Commit();
      }
    }

    private static void ReportPlan(Editor ed, CircuitMap map, RemapPlan plan)
    {
      foreach (KeyValuePair<int, int> pair in map.Pairs)
      {
        plan.Counts.TryGetValue(pair.Key, out int count);
        string note = count == 0 ? "  (not found)" : string.Empty;
        ed.WriteMessage($"\n  {map.OldPanel}-{pair.Key}: {count} label(s){note}");
      }
      foreach (string partial in plan.Partials.Distinct(StringComparer.OrdinalIgnoreCase))
      {
        ed.WriteMessage($"\n  Skipped \"{partial}\": only some of its circuits are in the range. Change it by hand.");
      }
      if (plan.LockedCount > 0)
      {
        ed.WriteMessage($"\n  Skipped {plan.LockedCount} object(s) on locked layers.");
      }
      if (plan.Edits.Count == 0)
      {
        ed.WriteMessage("\nNo labels to update.");
      }
    }

    private static HashSet<ObjectId> FindLockedLayers(Transaction tr, Database db)
    {
      HashSet<ObjectId> locked = new HashSet<ObjectId>();
      LayerTable layers = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
      foreach (ObjectId layerId in layers)
      {
        LayerTableRecord layer = tr.GetObject(layerId, OpenMode.ForRead, false) as LayerTableRecord;
        if (layer != null && layer.IsLocked)
        {
          locked.Add(layerId);
        }
      }
      return locked;
    }

    private static IEnumerable<ObjectId> EnumerateLayoutEntities(Transaction tr, Database db)
    {
      List<ObjectId> ids = new List<ObjectId>();
      BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
      foreach (ObjectId blockId in blockTable)
      {
        BlockTableRecord block = tr.GetObject(blockId, OpenMode.ForRead, false) as BlockTableRecord;
        if (block == null || !block.IsLayout)
        {
          continue;
        }
        foreach (ObjectId entityId in block)
        {
          ids.Add(entityId);
        }
      }
      return ids;
    }

    private enum TextKind
    {
      Text,
      MText,
      MTextAttribute,
      Leader,
    }

    private sealed class RemapEdit
    {
      internal ObjectId Id { get; set; }
      internal TextKind Kind { get; set; }
      internal string Contents { get; set; }
    }

    private sealed class RemapPlan
    {
      internal List<RemapEdit> Edits { get; } = new List<RemapEdit>();
      internal Dictionary<int, int> Counts { get; } = new Dictionary<int, int>();
      internal List<string> Partials { get; } = new List<string>();
      internal int LockedCount { get; set; }
    }

    internal sealed class CircuitMap
    {
      private readonly SortedDictionary<int, int> _pairs = new SortedDictionary<int, int>();
      private readonly Regex _labelPattern;

      internal CircuitMap(string oldPanel, int oldFirst, int oldLast, int step, string newPanel, int newFirst)
      {
        OldPanel = NormalizePanel(oldPanel);
        NewPanel = NormalizePanel(newPanel);
        for (int circuit = oldFirst; circuit <= oldLast; circuit += step)
        {
          _pairs[circuit] = newFirst + (circuit - oldFirst);
        }

        // The panel name must not be glued to a preceding letter/digit (so DP- never matches inside
        // XDP-), except an MText code like \P or \L, which ends in a letter but is not label text.
        // Multi-pole labels list their circuits with commas or ampersands: DP-16,18,20 or DP-16&18.
        _labelPattern = new Regex(
          @"(?<!(?<!\\)[A-Za-z0-9])" + Regex.Escape(OldPanel) + @"-(?<circuits>\d+(?:[ ]*[,&][ ]*\d+)*)(?!\d)",
          RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
      }

      internal string OldPanel { get; }
      internal string NewPanel { get; }
      internal int Count => _pairs.Count;
      internal IEnumerable<KeyValuePair<int, int>> Pairs => _pairs;

      internal string Apply(string contents, IDictionary<int, int> counts, out List<string> partials)
      {
        List<string> skipped = new List<string>();
        partials = skipped;
        if (string.IsNullOrEmpty(contents))
        {
          return contents;
        }

        return _labelPattern.Replace(contents, match =>
        {
          System.Text.RegularExpressions.Group circuits = match.Groups["circuits"];
          MatchCollection numbers = Regex.Matches(circuits.Value, @"\d+");
          List<int> oldCircuits = numbers.Cast<Match>().Select(number => int.Parse(number.Value)).ToList();
          int mapped = oldCircuits.Count(_pairs.ContainsKey);
          if (mapped == 0)
          {
            return match.Value;
          }
          if (mapped < oldCircuits.Count)
          {
            skipped.Add(match.Value);
            return match.Value;
          }

          foreach (int circuit in oldCircuits)
          {
            counts.TryGetValue(circuit, out int count);
            counts[circuit] = count + 1;
          }
          string newCircuits = Regex.Replace(circuits.Value, @"\d+", number => _pairs[int.Parse(number.Value)].ToString());
          return NewPanel + "-" + newCircuits;
        });
      }
    }
  }
}
