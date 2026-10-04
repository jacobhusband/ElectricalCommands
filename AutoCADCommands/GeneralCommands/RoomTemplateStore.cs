using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Newtonsoft.Json;

namespace ElectricalCommands
{
  /// <summary>
  /// Room templates saved in the drawing's Named Objects Dictionary so they travel
  /// with the DWG. Each template is one Xrecord holding its JSON in short chunks,
  /// because a single DWG text value cannot hold a whole template.
  /// </summary>
  internal static class RoomTemplateStore
  {
    private const string DictionaryKey = "ACIES_ROOM_TEMPLATES";
    private const int RecordVersion = 1;
    private const int ChunkLength = 1000;

    internal static void Save(Database database, RoomTemplate template)
    {
      string json = JsonConvert.SerializeObject(template, Formatting.None);
      List<TypedValue> values = new List<TypedValue>
      {
        new TypedValue((int)DxfCode.Int32, RecordVersion),
      };
      for (int start = 0; start < json.Length; start += ChunkLength)
      {
        values.Add(
          new TypedValue(
            (int)DxfCode.Text,
            json.Substring(start, Math.Min(ChunkLength, json.Length - start))
          )
        );
      }

      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        DBDictionary templates = OpenDictionary(transaction, database, true);
        string key = BuildKey(template.Name);
        ResultBuffer data = new ResultBuffer(values.ToArray());
        if (templates.Contains(key))
        {
          Xrecord existing = (Xrecord)transaction.GetObject(
            templates.GetAt(key),
            OpenMode.ForWrite
          );
          existing.Data = data;
        }
        else
        {
          Xrecord record = new Xrecord { Data = data };
          templates.SetAt(key, record);
          transaction.AddNewlyCreatedDBObject(record, true);
        }
        transaction.Commit();
      }
    }

    internal static bool TryLoad(Database database, string name, out RoomTemplate template)
    {
      template = null;
      foreach (RoomTemplate candidate in ReadAll(database))
      {
        if (string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
        {
          template = candidate;
          return true;
        }
      }
      return false;
    }

    internal static List<string> ListNames(Database database)
    {
      List<string> names = new List<string>();
      foreach (RoomTemplate template in ReadAll(database))
      {
        names.Add(template.Name);
      }
      names.Sort(StringComparer.OrdinalIgnoreCase);
      return names;
    }

    internal static bool Delete(Database database, string name)
    {
      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        DBDictionary templates = OpenDictionary(transaction, database, false);
        string key = BuildKey(name);
        if (templates == null || !templates.Contains(key))
        {
          return false;
        }

        templates.UpgradeOpen();
        DBObject record = transaction.GetObject(templates.GetAt(key), OpenMode.ForWrite);
        record.Erase();
        transaction.Commit();
        return true;
      }
    }

    private static List<RoomTemplate> ReadAll(Database database)
    {
      List<RoomTemplate> all = new List<RoomTemplate>();
      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        DBDictionary templates = OpenDictionary(transaction, database, false);
        if (templates == null)
        {
          return all;
        }

        foreach (DBDictionaryEntry entry in templates)
        {
          Xrecord record = transaction.GetObject(entry.Value, OpenMode.ForRead, false) as Xrecord;
          TypedValue[] values = record?.Data?.AsArray();
          if (values == null || values.Length < 2 || Convert.ToInt32(values[0].Value) != RecordVersion)
          {
            continue;
          }

          StringBuilder json = new StringBuilder();
          for (int index = 1; index < values.Length; index++)
          {
            json.Append(Convert.ToString(values[index].Value));
          }

          try
          {
            RoomTemplate template = JsonConvert.DeserializeObject<RoomTemplate>(json.ToString());
            if (template != null && !string.IsNullOrWhiteSpace(template.Name))
            {
              all.Add(template);
            }
          }
          catch (JsonException)
          {
            // A damaged template is skipped rather than blocking the others.
          }
        }
      }
      return all;
    }

    private static DBDictionary OpenDictionary(
      Transaction transaction,
      Database database,
      bool createIfMissing
    )
    {
      DBDictionary namedObjects = (DBDictionary)transaction.GetObject(
        database.NamedObjectsDictionaryId,
        createIfMissing ? OpenMode.ForWrite : OpenMode.ForRead,
        false
      );
      if (namedObjects.Contains(DictionaryKey))
      {
        return (DBDictionary)transaction.GetObject(
          namedObjects.GetAt(DictionaryKey),
          createIfMissing ? OpenMode.ForWrite : OpenMode.ForRead,
          false
        );
      }
      if (!createIfMissing)
      {
        return null;
      }

      if (!namedObjects.IsWriteEnabled)
      {
        namedObjects.UpgradeOpen();
      }
      DBDictionary created = new DBDictionary();
      namedObjects.SetAt(DictionaryKey, created);
      transaction.AddNewlyCreatedDBObject(created, true);
      return created;
    }

    // Dictionary keys cannot hold several punctuation characters, so the template's
    // real name lives in its JSON and the key is only a sanitized, case-insensitive handle.
    private static string BuildKey(string name)
    {
      StringBuilder key = new StringBuilder();
      foreach (char letter in (name ?? string.Empty).Trim().ToUpperInvariant())
      {
        key.Append(char.IsLetterOrDigit(letter) || letter == '-' || letter == '_' ? letter : '_');
      }
      return key.Length == 0 ? "UNNAMED" : key.ToString();
    }
  }
}
