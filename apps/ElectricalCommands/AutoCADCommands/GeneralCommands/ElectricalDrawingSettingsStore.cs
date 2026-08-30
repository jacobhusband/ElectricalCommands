using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;

namespace ElectricalCommands
{
  /// <summary>
  /// Drawing-resident settings shared by electrical drafting commands.
  /// Values are stored in the Named Objects Dictionary so they travel with the DWG.
  /// </summary>
  internal static class ElectricalDrawingSettingsStore
  {
    private const string DictionaryKey = "ACIES_ELECTRICAL_COMMAND_SETTINGS";
    private const string PanelLocationKey = "PANEL_LOCATION";
    private const string ScaleKey = "DRAWING_SCALE";
    private const string PanelNameKey = "PANEL_NAME";
    private const string PanelScheduleKey = "PANEL_SCHEDULE";
    private const string HomerunLayerKey = "HOMERUN_LAYER";
    private const string ReceptacleCircuitMaxKvaKey =
      "RECEPTACLE_CIRCUIT_MAX_KVA";
    private const int RecordVersion = 1;

    internal sealed class PanelLocationSetting
    {
      public Point3d Point { get; set; }
      public string SpaceHandle { get; set; } = string.Empty;
      public string Context { get; set; } = string.Empty;
    }

    internal sealed class ScaleSetting
    {
      /// <summary>Paper inches representing one model foot (0.25 for 1/4" = 1'-0").</summary>
      public double PaperInchesPerModelFoot { get; set; }
      public string DisplayText { get; set; } = string.Empty;
    }

    internal sealed class PanelScheduleSetting
    {
      public string WorkbookPath { get; set; } = string.Empty;
      public int CircuitCapacity { get; set; }
    }

    public static void WritePanelLocation(
      Database database,
      Point3d point,
      string spaceHandle,
      string context
    )
    {
      WriteRecord(
        database,
        PanelLocationKey,
        new ResultBuffer(
          new TypedValue((int)DxfCode.Int32, RecordVersion),
          new TypedValue((int)DxfCode.Real, point.X),
          new TypedValue((int)DxfCode.Real, point.Y),
          new TypedValue((int)DxfCode.Real, point.Z),
          new TypedValue((int)DxfCode.Text, spaceHandle ?? string.Empty),
          new TypedValue((int)DxfCode.Text, context ?? string.Empty)
        )
      );
    }

    public static bool TryReadPanelLocation(
      Database database,
      out PanelLocationSetting setting
    )
    {
      setting = null;
      TypedValue[] values = ReadRecord(database, PanelLocationKey);
      if (values == null || values.Length < 6 || !HasSupportedVersion(values))
      {
        return false;
      }

      try
      {
        setting = new PanelLocationSetting
        {
          Point = new Point3d(
            Convert.ToDouble(values[1].Value),
            Convert.ToDouble(values[2].Value),
            Convert.ToDouble(values[3].Value)
          ),
          SpaceHandle = Convert.ToString(values[4].Value) ?? string.Empty,
          Context = Convert.ToString(values[5].Value) ?? string.Empty,
        };
        return true;
      }
      catch
      {
        setting = null;
        return false;
      }
    }

    public static void WriteScale(
      Database database,
      double paperInchesPerModelFoot,
      string displayText
    )
    {
      WriteRecord(
        database,
        ScaleKey,
        new ResultBuffer(
          new TypedValue((int)DxfCode.Int32, RecordVersion),
          new TypedValue((int)DxfCode.Real, paperInchesPerModelFoot),
          new TypedValue((int)DxfCode.Text, displayText ?? string.Empty)
        )
      );
    }

    public static bool TryReadScale(Database database, out ScaleSetting setting)
    {
      setting = null;
      TypedValue[] values = ReadRecord(database, ScaleKey);
      if (values == null || values.Length < 3 || !HasSupportedVersion(values))
      {
        return false;
      }

      try
      {
        double value = Convert.ToDouble(values[1].Value);
        if (value <= 0.0 || double.IsNaN(value) || double.IsInfinity(value))
        {
          return false;
        }

        setting = new ScaleSetting
        {
          PaperInchesPerModelFoot = value,
          DisplayText = Convert.ToString(values[2].Value) ?? string.Empty,
        };
        return true;
      }
      catch
      {
        setting = null;
        return false;
      }
    }

    public static void WritePanelName(Database database, string panelName)
    {
      WriteRecord(
        database,
        PanelNameKey,
        new ResultBuffer(
          new TypedValue((int)DxfCode.Int32, RecordVersion),
          new TypedValue((int)DxfCode.Text, panelName ?? string.Empty)
        )
      );
    }

    public static bool TryReadPanelName(Database database, out string panelName)
    {
      panelName = string.Empty;
      TypedValue[] values = ReadRecord(database, PanelNameKey);
      if (values == null || values.Length < 2 || !HasSupportedVersion(values))
      {
        return false;
      }

      panelName = (Convert.ToString(values[1].Value) ?? string.Empty).Trim();
      return panelName.Length > 0;
    }

    public static void WritePanelSchedule(
      Database database,
      string workbookPath,
      int circuitCapacity)
    {
      WriteRecord(
        database,
        PanelScheduleKey,
        new ResultBuffer(
          new TypedValue((int)DxfCode.Int32, RecordVersion),
          new TypedValue((int)DxfCode.Text, workbookPath ?? string.Empty),
          new TypedValue((int)DxfCode.Int32, circuitCapacity)
        )
      );
    }

    public static bool TryReadPanelSchedule(
      Database database,
      out PanelScheduleSetting setting)
    {
      setting = null;
      TypedValue[] values = ReadRecord(database, PanelScheduleKey);
      if (values == null || values.Length < 3 || !HasSupportedVersion(values))
      {
        return false;
      }

      try
      {
        string workbookPath =
          (Convert.ToString(values[1].Value) ?? string.Empty).Trim();
        int circuitCapacity = Convert.ToInt32(values[2].Value);
        if (workbookPath.Length == 0 ||
            circuitCapacity < 6 ||
            circuitCapacity % 2 != 0)
        {
          return false;
        }

        setting = new PanelScheduleSetting
        {
          WorkbookPath = workbookPath,
          CircuitCapacity = circuitCapacity,
        };
        return true;
      }
      catch
      {
        setting = null;
        return false;
      }
    }

    public static void WriteReceptacleCircuitMaxKva(
      Database database,
      double maximumKva)
    {
      if (!TryNormalizeReceptacleCircuitMaxKva(
        maximumKva,
        out double normalizedKva))
      {
        throw new ArgumentOutOfRangeException(
          nameof(maximumKva),
          $"Maximum receptacle circuit load must be a multiple of " +
          $"{GeneralCommands.ReceptacleLoadUnitKva:0.00} kVA between " +
          $"{GeneralCommands.ReceptacleLoadUnitKva:0.00} and " +
          $"{GeneralCommands.MaximumReceptacleCircuitLoadKva:0.00} kVA.");
      }

      WriteRecord(
        database,
        ReceptacleCircuitMaxKvaKey,
        new ResultBuffer(
          new TypedValue((int)DxfCode.Int32, RecordVersion),
          new TypedValue((int)DxfCode.Real, normalizedKva)
        )
      );
    }

    public static bool TryReadReceptacleCircuitMaxKva(
      Database database,
      out double maximumKva)
    {
      maximumKva = 0.0;
      TypedValue[] values = ReadRecord(
        database,
        ReceptacleCircuitMaxKvaKey);
      if (values == null || values.Length < 2 || !HasSupportedVersion(values))
      {
        return false;
      }

      try
      {
        return TryNormalizeReceptacleCircuitMaxKva(
          Convert.ToDouble(values[1].Value),
          out maximumKva);
      }
      catch
      {
        maximumKva = 0.0;
        return false;
      }
    }

    private static bool TryNormalizeReceptacleCircuitMaxKva(
      double maximumKva,
      out double normalizedKva)
    {
      normalizedKva = 0.0;
      if (maximumKva <= 0.0 ||
          double.IsNaN(maximumKva) ||
          double.IsInfinity(maximumKva))
      {
        return false;
      }

      int loadUnits = (int)Math.Round(
        maximumKva / GeneralCommands.ReceptacleLoadUnitKva);
      if (loadUnits < 1 ||
          loadUnits > GeneralCommands.MaximumReceptacleCircuitLoadUnits)
      {
        return false;
      }

      normalizedKva =
        loadUnits * GeneralCommands.ReceptacleLoadUnitKva;
      return Math.Abs(maximumKva - normalizedKva) < 0.001;
    }

    public static void WriteHomerunLayer(Database database, string layerName)
    {
      WriteRecord(
        database,
        HomerunLayerKey,
        new ResultBuffer(
          new TypedValue((int)DxfCode.Int32, RecordVersion),
          new TypedValue((int)DxfCode.Text, layerName ?? string.Empty)
        )
      );
    }

    public static bool TryReadHomerunLayer(Database database, out string layerName)
    {
      layerName = string.Empty;
      TypedValue[] values = ReadRecord(database, HomerunLayerKey);
      if (values == null || values.Length < 2 || !HasSupportedVersion(values))
      {
        return false;
      }

      layerName = (Convert.ToString(values[1].Value) ?? string.Empty).Trim();
      return layerName.Length > 0;
    }

    private static bool HasSupportedVersion(TypedValue[] values)
    {
      try
      {
        return Convert.ToInt32(values[0].Value) == RecordVersion;
      }
      catch
      {
        return false;
      }
    }

    private static TypedValue[] ReadRecord(Database database, string recordKey)
    {
      try
      {
        using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
        {
          DBDictionary namedObjects = transaction.GetObject(
            database.NamedObjectsDictionaryId,
            OpenMode.ForRead,
            false
          ) as DBDictionary;
          if (namedObjects == null || !namedObjects.Contains(DictionaryKey))
          {
            return null;
          }

          DBDictionary settings = transaction.GetObject(
            namedObjects.GetAt(DictionaryKey),
            OpenMode.ForRead,
            false
          ) as DBDictionary;
          if (settings == null || !settings.Contains(recordKey))
          {
            return null;
          }

          Xrecord record = transaction.GetObject(
            settings.GetAt(recordKey),
            OpenMode.ForRead,
            false
          ) as Xrecord;
          return record?.Data?.AsArray();
        }
      }
      catch
      {
        return null;
      }
    }

    private static void WriteRecord(Database database, string recordKey, ResultBuffer data)
    {
      using (Transaction transaction = database.TransactionManager.StartTransaction())
      {
        DBDictionary namedObjects = transaction.GetObject(
          database.NamedObjectsDictionaryId,
          OpenMode.ForWrite
        ) as DBDictionary;
        if (namedObjects == null)
        {
          throw new InvalidOperationException("Unable to open the drawing settings dictionary.");
        }

        DBDictionary settings;
        if (namedObjects.Contains(DictionaryKey))
        {
          settings = transaction.GetObject(
            namedObjects.GetAt(DictionaryKey),
            OpenMode.ForWrite
          ) as DBDictionary;
        }
        else
        {
          settings = new DBDictionary();
          namedObjects.SetAt(DictionaryKey, settings);
          transaction.AddNewlyCreatedDBObject(settings, true);
        }

        if (settings == null)
        {
          throw new InvalidOperationException("Unable to create the electrical command settings dictionary.");
        }

        if (settings.Contains(recordKey))
        {
          Xrecord existing = transaction.GetObject(
            settings.GetAt(recordKey),
            OpenMode.ForWrite
          ) as Xrecord;
          if (existing == null)
          {
            throw new InvalidOperationException("The drawing setting record is invalid.");
          }
          existing.Data = data;
        }
        else
        {
          Xrecord record = new Xrecord { Data = data };
          settings.SetAt(recordKey, record);
          transaction.AddNewlyCreatedDBObject(record, true);
        }

        transaction.Commit();
      }
    }
  }
}
