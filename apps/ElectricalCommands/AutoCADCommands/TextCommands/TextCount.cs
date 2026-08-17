using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ElectricalCommands
{
  public static class TextCountCommand
  {
    [CommandMethod("TEXTCOUNT", CommandFlags.UsePickSet)]
    public static void TextCount()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      try
      {
        PromptSelectionResult selection = ed.SelectImplied();
        if (selection.Status != PromptStatus.OK)
        {
          PromptSelectionOptions options = new PromptSelectionOptions();
          options.MessageForAdding = "\nSelect TEXT, MTEXT, or ATTRIB objects: ";
          TypedValue[] filterValues =
          {
            new TypedValue((int)DxfCode.Start, "TEXT,MTEXT,ATTRIB")
          };
          SelectionFilter filter = new SelectionFilter(filterValues);
          selection = ed.GetSelection(options, filter);
        }

        if (selection.Status != PromptStatus.OK || selection.Value == null || selection.Value.Count == 0)
        {
          ed.WriteMessage("\nNo TEXT, MTEXT, or ATTRIB objects selected.");
          return;
        }

        Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal);
        int total = 0;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId objectId in selection.Value.GetObjectIds())
          {
            DBObject obj = tr.GetObject(objectId, OpenMode.ForRead, false);
            string text = GetTextValue(obj);
            if (string.IsNullOrWhiteSpace(text))
            {
              continue;
            }

            total++;
            if (counts.ContainsKey(text))
            {
              counts[text]++;
            }
            else
            {
              counts[text] = 1;
            }
          }

          tr.Commit();
        }

        if (counts.Count == 0)
        {
          ed.WriteMessage("\nNo TEXT, MTEXT, or ATTRIB objects selected.");
          return;
        }

        List<KeyValuePair<string, int>> orderedCounts = counts
          .OrderByDescending(pair => pair.Value)
          .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
          .ToList();

        ed.WriteMessage("\n--- Text Object Count ---");
        foreach (KeyValuePair<string, int> pair in orderedCounts)
        {
          ed.WriteMessage($"\n{pair.Key} : {pair.Value}");
        }
        ed.WriteMessage(
          $"\n-------------------------\nTotal unique text objects: {orderedCounts.Count}\nTotal objects processed: {total}"
        );

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (!string.IsNullOrWhiteSpace(desktopPath) && Directory.Exists(desktopPath))
        {
          string outputPath = Path.Combine(
            desktopPath,
            $"TextCount_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.txt"
          );
          WriteTextCountReport(outputPath, orderedCounts, total);

          try
          {
            Process.Start(new ProcessStartInfo
            {
              FileName = "notepad.exe",
              Arguments = $"\"{outputPath}\"",
              UseShellExecute = true
            });
            ed.WriteMessage("\n\nExported counts to TXT and opened in Notepad.");
          }
          catch
          {
            ed.WriteMessage($"\n\nWrote counts to: {outputPath}");
          }
        }
        else
        {
          ed.WriteMessage("\n\nCould not determine desktop path.");
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nTEXTCOUNT error: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(Array.Empty<ObjectId>());
      }
    }

    private static string GetTextValue(DBObject obj)
    {
      if (obj is AttributeReference attribute)
      {
        return NormalizeText(attribute.TextString);
      }

      if (obj is DBText text)
      {
        return NormalizeText(text.TextString);
      }

      if (obj is MText mText)
      {
        return NormalizeText(mText.Text);
      }

      return string.Empty;
    }

    private static string NormalizeText(string value)
    {
      return string.IsNullOrWhiteSpace(value)
        ? string.Empty
        : value.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private static void WriteTextCountReport(
      string outputPath,
      IReadOnlyList<KeyValuePair<string, int>> counts,
      int total
    )
    {
      using (StreamWriter writer = new StreamWriter(outputPath, false))
      {
        writer.WriteLine($"Total selected text objects: {total}");
        writer.WriteLine();
        foreach (KeyValuePair<string, int> pair in counts)
        {
          writer.WriteLine($"{pair.Key} : {pair.Value}");
        }
      }
    }
  }
}
