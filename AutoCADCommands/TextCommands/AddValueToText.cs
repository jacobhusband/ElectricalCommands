using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("TXTADD", CommandFlags.UsePickSet)]
    [CommandMethod("TA", CommandFlags.UsePickSet)]
    public void Add2Txt()
    {
      Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      try
      {
        SelectionSet sset;
        PromptSelectionResult selRes = ed.SelectImplied();

        if (selRes.Status == PromptStatus.OK)
        {
          // Use the PickFirst selection
          sset = selRes.Value;
        }
        else
        {
          // If no PickFirst selection, prompt for selection
          PromptSelectionOptions pso = new PromptSelectionOptions();
          pso.MessageForAdding = "Select DBText and MText objects: ";
          TypedValue[] filterList = new TypedValue[]
          {
                new TypedValue((int)DxfCode.Start, "TEXT,MTEXT")
          };
          SelectionFilter filter = new SelectionFilter(filterList);
          selRes = ed.GetSelection(pso, filter);
          if (selRes.Status != PromptStatus.OK)
            return;
          sset = selRes.Value;
        }

        // Filter for TEXT and MTEXT objects
        ObjectId[] filteredIds;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          filteredIds = sset.GetObjectIds().Where(id =>
          {
            var obj = tr.GetObject(id, OpenMode.ForRead);
            return obj is DBText || obj is MText;
          }).ToArray();
          tr.Commit();
        }

        if (filteredIds.Length == 0)
        {
          ed.WriteMessage("\nNo text objects selected.");
          return;
        }

        // Prompt user for the increment value
        PromptIntegerOptions pio = new PromptIntegerOptions("Enter the value to add: ");
        PromptIntegerResult intRes = ed.GetInteger(pio);
        if (intRes.Status != PromptStatus.OK)
          return;
        int incrementValue = intRes.Value;

        // Process selected objects
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          int processedCount = 0;
          foreach (ObjectId objId in filteredIds)
          {
            Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
            if (ent is DBText text)
            {
              text.TextString = ProcessText(RemoveStyling(text.TextString), incrementValue);
              processedCount++;
            }
            else if (ent is MText mtext)
            {
              mtext.Contents = ProcessText(RemoveStyling(mtext.Contents), incrementValue);
              processedCount++;
            }
          }
          tr.Commit();
          ed.WriteMessage($"\nCommand completed successfully. Processed {processedCount} text object(s).");
        }

        // Clear the PickFirst selection set
        ed.SetImpliedSelection(new ObjectId[0]);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("\nError: " + ex.Message);
      }
    }

    private string RemoveStyling(string input)
    {
      string withoutBrackets = Regex.Replace(input, @"\{[^}]*\}", "");

      int semicolonIndex = withoutBrackets.IndexOf(';');
      if (semicolonIndex != -1)
      {
        withoutBrackets = withoutBrackets.Substring(semicolonIndex + 1);
      }

      return withoutBrackets.Trim();
    }

    private string ProcessText(string input, int increment)
    {
      // Case 1: Text matches the pattern "LB-1&3&5" (prefix with '&' separated values)
      if (Regex.IsMatch(input, @"^[A-Za-z]+-\d+(&\d+)*$"))
      {
        string[] parts = input.Split('-');
        string prefix = parts[0];
        string[] numbers = parts[1].Split('&');
        string result = string.Join("&", numbers.Select(n => (int.Parse(n) + increment).ToString()));
        return $"{prefix}-{result}";
      }
      // Case 2: Text matches the pattern "1&3&5" (only '&' separated values without prefix)
      else if (Regex.IsMatch(input, @"^\d+(&\d+)*$"))
      {
        string[] numbers = input.Split('&');
        return string.Join("&", numbers.Select(n => (int.Parse(n) + increment).ToString()));
      }
      // Case 3: Text matches the pattern "LB-1,3,5" (prefix with comma separated values)
      else if (Regex.IsMatch(input, @"^[A-Za-z]+-\d+(,\d+)*$"))
      {
        string[] parts = input.Split('-');
        string prefix = parts[0];
        string[] numbers = parts[1].Split(',');
        string result = string.Join(",", numbers.Select(n => (int.Parse(n) + increment).ToString()));
        return $"{prefix}-{result}";
      }
      // Case 4: Text matches the pattern "1,3,5" (only comma separated values without prefix)
      else if (Regex.IsMatch(input, @"^\d+(,\d+)*$"))
      {
        string[] numbers = input.Split(',');
        return string.Join(",", numbers.Select(n => (int.Parse(n) + increment).ToString()));
      }
      // Case 5: Text contains only a single number
      else if (int.TryParse(input, out int value))
      {
        return (value + increment).ToString();
      }
      // If none of the conditions are met, return the original input
      else
      {
        return input;
      }
    }
  }
}
