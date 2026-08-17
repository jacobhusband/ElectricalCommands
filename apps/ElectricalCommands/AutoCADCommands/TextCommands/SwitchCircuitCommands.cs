using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ElectricalCommands
{
  public sealed class SwitchCircuitCommands : IExtensionApplication
  {
    private const string StateDictionaryKey = "ELECTRICALCOMMANDS_SWITCHCIRCUIT";
    private const string NextIndexRecordKey = "NEXT_INDEX";
    private const string AssignmentRecordKey = "ELECTRICALCOMMANDS_SWITCHCIRCUIT_ASSIGNMENT";
    private const int AssignmentVersion = 1;

    public void Initialize()
    {
      AcApplication.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
    }

    public void Terminate()
    {
      AcApplication.DocumentManager.DocumentActivated -= DocumentManager_DocumentActivated;
      SwitchCircuitPalette.Dispose();
    }

    [CommandMethod("SWITCHCIRCUIT", CommandFlags.Modal)]
    [CommandMethod("SCIRCUIT", CommandFlags.Modal)]
    public void OpenPalette()
    {
      SwitchCircuitPalette.Show();
    }

    [CommandMethod("-SWITCHCIRCUITAPPLY", CommandFlags.UsePickSet | CommandFlags.NoHistory)]
    public void ApplySwitchCircuits()
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        return;
      }

      Editor editor = document.Editor;
      PromptSelectionResult selection = editor.SelectImplied();
      if (selection.Status != PromptStatus.OK || selection.Value == null || selection.Value.Count == 0)
      {
        PromptSelectionOptions options = new PromptSelectionOptions
        {
          MessageForAdding = "\nSelect fixture-tag MText objects: ",
        };
        SelectionFilter filter = new SelectionFilter(new[]
        {
          new TypedValue((int)DxfCode.Start, "MTEXT"),
        });
        selection = editor.GetSelection(options, filter);
      }

      if (selection.Status != PromptStatus.OK || selection.Value == null || selection.Value.Count == 0)
      {
        SetResult(editor, "Switch circuit assignment canceled.");
        return;
      }

      try
      {
        ApplySelection(document.Database, selection.Value.GetObjectIds(), out SwitchCircuitApplyResult result);
        string message = BuildApplyMessage(result);
        editor.WriteMessage("\n" + message);
        SwitchCircuitPalette.SetStatus(message);
        SwitchCircuitPalette.Refresh();
      }
      catch (System.Exception ex)
      {
        SetResult(editor, "SWITCHCIRCUIT error: " + ex.Message);
      }
      finally
      {
        editor.SetImpliedSelection(Array.Empty<ObjectId>());
      }
    }

    internal static bool TrySetNextIndex(
      Database database,
      int requestedIndex,
      out string message
    )
    {
      if (database == null)
      {
        message = "No active drawing is available.";
        return false;
      }

      if (requestedIndex < 0)
      {
        message = "The next circuit index is invalid.";
        return false;
      }

      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        WriteStoredNextIndex(transaction, database, requestedIndex);
        transaction.Commit();
      }

      message = $"Next switch circuit set to {SwitchCircuitSequence.ToSuffix(requestedIndex)}.";
      return true;
    }

    internal static int GetEffectiveNextIndex(Database database)
    {
      if (database == null)
      {
        return 0;
      }

      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        return TryReadStoredNextIndex(transaction, database, out int storedIndex)
          ? storedIndex
          : FindAssignmentHighWater(transaction, database);
      }
    }

    private static void ApplySelection(
      Database database,
      IEnumerable<ObjectId> selectedIds,
      out SwitchCircuitApplyResult result
    )
    {
      result = new SwitchCircuitApplyResult();
      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        Dictionary<string, List<ObjectId>> groups = new Dictionary<string, List<ObjectId>>(StringComparer.Ordinal);
        foreach (ObjectId objectId in selectedIds ?? Enumerable.Empty<ObjectId>())
        {
          DBObject selectedObject;
          try
          {
            selectedObject = transaction.GetObject(objectId, OpenMode.ForRead, false);
          }
          catch
          {
            result.IgnoredCount++;
            continue;
          }

          MText mtext = selectedObject as MText;
          if (mtext == null)
          {
            result.IgnoredCount++;
            continue;
          }

          if (TryReadAssignmentSuffix(transaction, mtext, out _))
          {
            result.SkippedCount++;
            continue;
          }

          string fixtureKey = NormalizeVisibleText(mtext.Text);
          if (!groups.TryGetValue(fixtureKey, out List<ObjectId> group))
          {
            group = new List<ObjectId>();
            groups.Add(fixtureKey, group);
          }
          group.Add(objectId);
        }

        if (groups.Count == 0)
        {
          return;
        }

        int nextIndex = TryReadStoredNextIndex(transaction, database, out int storedIndex)
          ? storedIndex
          : FindAssignmentHighWater(transaction, database);
        foreach (KeyValuePair<string, List<ObjectId>> group in groups.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
          string suffix = SwitchCircuitSequence.ToSuffix(nextIndex);
          foreach (ObjectId objectId in group.Value)
          {
            MText mtext = transaction.GetObject(objectId, OpenMode.ForWrite, false) as MText;
            if (mtext == null)
            {
              throw new InvalidOperationException("A selected MText object became unavailable.");
            }

            mtext.Contents = (mtext.Contents ?? string.Empty) + suffix;
            WriteAssignmentMarker(transaction, mtext, group.Key, suffix);
            result.ModifiedCount++;
          }

          result.GroupCount++;
          nextIndex = checked(nextIndex + 1);
        }

        WriteStoredNextIndex(transaction, database, nextIndex);
        result.NextIndex = nextIndex;
        transaction.Commit();
      }
    }

    private static int FindAssignmentHighWater(Transaction transaction, Database database)
    {
      int highWater = 0;
      BlockTable blockTable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
      if (blockTable == null)
      {
        return highWater;
      }

      foreach (ObjectId blockRecordId in blockTable)
      {
        BlockTableRecord blockRecord;
        try
        {
          blockRecord = transaction.GetObject(blockRecordId, OpenMode.ForRead, false) as BlockTableRecord;
        }
        catch
        {
          continue;
        }
        if (blockRecord == null)
        {
          continue;
        }

        foreach (ObjectId entityId in blockRecord)
        {
          MText mtext;
          try
          {
            mtext = transaction.GetObject(entityId, OpenMode.ForRead, false) as MText;
          }
          catch
          {
            continue;
          }
          if (
            mtext != null &&
            TryReadAssignmentSuffix(transaction, mtext, out string suffix) &&
            SwitchCircuitSequence.TryParseSuffix(suffix, out int assignedIndex)
          )
          {
            highWater = Math.Max(highWater, checked(assignedIndex + 1));
          }
        }
      }

      return highWater;
    }

    private static bool TryReadAssignmentSuffix(Transaction transaction, MText mtext, out string suffix)
    {
      suffix = string.Empty;
      try
      {
        if (mtext == null || mtext.ExtensionDictionary.IsNull)
        {
          return false;
        }

        DBDictionary dictionary = transaction.GetObject(
          mtext.ExtensionDictionary,
          OpenMode.ForRead,
          false
        ) as DBDictionary;
        if (dictionary == null || !dictionary.Contains(AssignmentRecordKey))
        {
          return false;
        }

        Xrecord record = transaction.GetObject(
          dictionary.GetAt(AssignmentRecordKey),
          OpenMode.ForRead,
          false
        ) as Xrecord;
        TypedValue[] values = record?.Data?.AsArray();
        if (values == null || values.Length < 3)
        {
          return false;
        }

        if (Convert.ToInt32(values[0].Value) != AssignmentVersion)
        {
          return false;
        }

        suffix = Convert.ToString(values[2].Value) ?? string.Empty;
        return SwitchCircuitSequence.TryParseSuffix(suffix, out _);
      }
      catch
      {
        suffix = string.Empty;
        return false;
      }
    }

    private static void WriteAssignmentMarker(
      Transaction transaction,
      MText mtext,
      string fixtureKey,
      string suffix
    )
    {
      if (mtext.ExtensionDictionary.IsNull)
      {
        mtext.CreateExtensionDictionary();
      }

      DBDictionary dictionary = transaction.GetObject(
        mtext.ExtensionDictionary,
        OpenMode.ForWrite
      ) as DBDictionary;
      if (dictionary == null)
      {
        throw new InvalidOperationException("Unable to create MText assignment metadata.");
      }

      Xrecord record = new Xrecord
      {
        Data = new ResultBuffer(
          new TypedValue((int)DxfCode.Int32, AssignmentVersion),
          new TypedValue((int)DxfCode.Text, fixtureKey),
          new TypedValue((int)DxfCode.Text, suffix)
        ),
      };
      dictionary.SetAt(AssignmentRecordKey, record);
      transaction.AddNewlyCreatedDBObject(record, true);
    }

    private static bool TryReadStoredNextIndex(
      Transaction transaction,
      Database database,
      out int nextIndex
    )
    {
      nextIndex = 0;
      DBDictionary namedObjects = transaction.GetObject(
        database.NamedObjectsDictionaryId,
        OpenMode.ForRead
      ) as DBDictionary;
      if (namedObjects == null || !namedObjects.Contains(StateDictionaryKey))
      {
        return false;
      }

      DBDictionary stateDictionary = transaction.GetObject(
        namedObjects.GetAt(StateDictionaryKey),
        OpenMode.ForRead,
        false
      ) as DBDictionary;
      if (stateDictionary == null || !stateDictionary.Contains(NextIndexRecordKey))
      {
        return false;
      }

      Xrecord record = transaction.GetObject(
        stateDictionary.GetAt(NextIndexRecordKey),
        OpenMode.ForRead,
        false
      ) as Xrecord;
      TypedValue[] values = record?.Data?.AsArray();
      if (values == null || values.Length == 0)
      {
        return false;
      }

      int value = Convert.ToInt32(values[0].Value);
      nextIndex = Math.Max(0, value);
      return true;
    }

    private static void WriteStoredNextIndex(Transaction transaction, Database database, int nextIndex)
    {
      DBDictionary namedObjects = transaction.GetObject(
        database.NamedObjectsDictionaryId,
        OpenMode.ForWrite
      ) as DBDictionary;
      if (namedObjects == null)
      {
        throw new InvalidOperationException("Unable to open the drawing state dictionary.");
      }

      DBDictionary stateDictionary;
      if (namedObjects.Contains(StateDictionaryKey))
      {
        stateDictionary = transaction.GetObject(
          namedObjects.GetAt(StateDictionaryKey),
          OpenMode.ForWrite
        ) as DBDictionary;
      }
      else
      {
        stateDictionary = new DBDictionary();
        namedObjects.SetAt(StateDictionaryKey, stateDictionary);
        transaction.AddNewlyCreatedDBObject(stateDictionary, true);
      }

      if (stateDictionary == null)
      {
        throw new InvalidOperationException("Unable to create the switch circuit state dictionary.");
      }

      ResultBuffer data = new ResultBuffer(new TypedValue((int)DxfCode.Int32, nextIndex));
      if (stateDictionary.Contains(NextIndexRecordKey))
      {
        Xrecord existing = transaction.GetObject(
          stateDictionary.GetAt(NextIndexRecordKey),
          OpenMode.ForWrite
        ) as Xrecord;
        if (existing == null)
        {
          throw new InvalidOperationException("The switch circuit state record is invalid.");
        }
        existing.Data = data;
      }
      else
      {
        Xrecord record = new Xrecord { Data = data };
        stateDictionary.SetAt(NextIndexRecordKey, record);
        transaction.AddNewlyCreatedDBObject(record, true);
      }
    }

    private static string NormalizeVisibleText(string value)
    {
      return string.IsNullOrWhiteSpace(value)
        ? string.Empty
        : value.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private static string BuildApplyMessage(SwitchCircuitApplyResult result)
    {
      if (result.GroupCount == 0)
      {
        return $"No unassigned MText was changed. Skipped {result.SkippedCount}, ignored {result.IgnoredCount}.";
      }

      string nextSuffix = SwitchCircuitSequence.ToSuffix(result.NextIndex);
      return $"Assigned {result.GroupCount} circuit(s) to {result.ModifiedCount} MText object(s). " +
        $"Skipped {result.SkippedCount}, ignored {result.IgnoredCount}. Next: {nextSuffix}.";
    }

    private static void SetResult(Editor editor, string message)
    {
      editor?.WriteMessage("\n" + message);
      SwitchCircuitPalette.SetStatus(message);
    }

    private static void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
    {
      SwitchCircuitPalette.Refresh();
    }

    private sealed class SwitchCircuitApplyResult
    {
      internal int GroupCount { get; set; }
      internal int ModifiedCount { get; set; }
      internal int SkippedCount { get; set; }
      internal int IgnoredCount { get; set; }
      internal int NextIndex { get; set; }
    }
  }
}
