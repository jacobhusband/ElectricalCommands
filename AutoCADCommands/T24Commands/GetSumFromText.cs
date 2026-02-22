using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Linq;

namespace ElectricalCommands
{
  public static class GetSumFromTextCommand
  {
    [CommandMethod("TXTSUM", CommandFlags.UsePickSet)]
    [CommandMethod("TS", CommandFlags.UsePickSet)]
    public static void SUMTEXT()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      PromptSelectionResult selection = ed.SelectImplied();
      if (selection.Status != PromptStatus.OK)
      {
        PromptSelectionOptions opts = new PromptSelectionOptions();
        opts.MessageForAdding = "Select text objects to sum: ";
        opts.AllowDuplicates = false;
        opts.RejectObjectsOnLockedLayers = true;
        selection = ed.GetSelection(opts);
        if (selection.Status != PromptStatus.OK)
          return;
      }
      double sum = 0.0;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        foreach (SelectedObject so in selection.Value)
        {
          DBText text = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;
          MText mtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;
          if (text != null)
          {
            double value;
            string textString = text
                .TextString.Replace("\\FArial;", "")
                .Replace("\\L", "")
                .Replace("\\Farial|c0;", "")
                .Replace("{", "")
                .Replace("}", "")
                .Replace("\\I", "")
                .Trim();
            textString = new string(
                textString.Where(c => char.IsDigit(c) || c == '.').ToArray()
            );
            if (textString.Contains("VA"))
            {
              string[] sections = textString.Split(
                  new string[] { "VA" },
                  StringSplitOptions.None
              );
              foreach (string section in sections)
              {
                double num = 0;
                if (double.TryParse(section, out num))
                {
                  sum += num;
                }
              }
            }
            if (Double.TryParse(textString, out value))
              sum += value;
          }
          else if (mtext != null)
          {
            double value;
            string mTextContents = mtext
                .Contents.Replace("\\FArial;", "")
                .Replace("\\L", "")
                .Replace("\\Farial|c0;", "")
                .Replace("{", "")
                .Replace("}", "")
                .Replace("\\I", "")
                .Trim();
            if (mTextContents.Contains("VA"))
            {
              string[] sections = mTextContents.Split(
                  new string[] { "VA" },
                  StringSplitOptions.None
              );
              foreach (string section in sections)
              {
                double num = 0;
                var newSection = new string(
                    section.Where(c => char.IsDigit(c) || c == '.').ToArray()
                );
                if (double.TryParse(newSection, out num))
                {
                  sum += num;
                }
              }
            }
            else if (Double.TryParse(mTextContents, out value))
            {
              sum += value;
            }
            else
            {
              mTextContents = new string(
                  mTextContents.Where(c => char.IsDigit(c) || c == '.').ToArray()
              );
              if (Double.TryParse(mTextContents, out value))
                sum += value;
            }
          }
        }
        ed.WriteMessage($"\nThe sum of selected text objects is: {sum}");
        tr.Commit();
      }
      // Clear the PickFirst selection set
      ed.SetImpliedSelection(new ObjectId[0]);
    }
  }
}
