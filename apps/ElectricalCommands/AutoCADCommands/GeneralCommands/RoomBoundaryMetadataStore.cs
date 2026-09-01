using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;

namespace ElectricalCommands
{
  /// <summary>
  /// Stores invisible room data directly on each source polyline so the data
  /// follows the entity when it is moved or copied within AutoCAD.
  /// </summary>
  internal static class RoomBoundaryMetadataStore
  {
    private const string RecordKey = "ACIES_ROOM_DATA";
    private const int RecordVersion = 1;

    internal sealed class RoomMetadata
    {
      internal string Name { get; set; } = string.Empty;
      internal double SquareFeet { get; set; }
      internal Point3d BasePoint { get; set; }
      internal Point2d RelativeLocation { get; set; }
    }

    internal static bool TryRead(
      Polyline polyline,
      Transaction transaction,
      out RoomMetadata metadata)
    {
      metadata = null;
      if (polyline == null ||
          transaction == null ||
          polyline.ExtensionDictionary.IsNull)
      {
        return false;
      }

      try
      {
        DBDictionary dictionary = transaction.GetObject(
          polyline.ExtensionDictionary,
          OpenMode.ForRead,
          false) as DBDictionary;
        if (dictionary == null || !dictionary.Contains(RecordKey))
        {
          return false;
        }

        Xrecord record = transaction.GetObject(
          dictionary.GetAt(RecordKey),
          OpenMode.ForRead,
          false) as Xrecord;
        TypedValue[] values = record?.Data?.AsArray();
        if (values == null || values.Length < 8 ||
            Convert.ToInt32(values[0].Value) != RecordVersion)
        {
          return false;
        }

        RoomMetadata result = new RoomMetadata
        {
          Name = (Convert.ToString(values[1].Value) ?? string.Empty).Trim(),
          SquareFeet = Convert.ToDouble(values[2].Value),
          BasePoint = new Point3d(
            Convert.ToDouble(values[3].Value),
            Convert.ToDouble(values[4].Value),
            Convert.ToDouble(values[5].Value)),
          RelativeLocation = new Point2d(
            Convert.ToDouble(values[6].Value),
            Convert.ToDouble(values[7].Value)),
        };
        if (result.Name.Length == 0 ||
            result.SquareFeet < 0.0 ||
            !IsFinite(result.SquareFeet) ||
            !IsFinite(result.BasePoint.X) ||
            !IsFinite(result.BasePoint.Y) ||
            !IsFinite(result.BasePoint.Z) ||
            !IsFinite(result.RelativeLocation.X) ||
            !IsFinite(result.RelativeLocation.Y))
        {
          return false;
        }

        metadata = result;
        return true;
      }
      catch
      {
        metadata = null;
        return false;
      }
    }

    internal static void Write(
      Polyline polyline,
      Transaction transaction,
      RoomMetadata metadata)
    {
      if (polyline == null)
      {
        throw new ArgumentNullException(nameof(polyline));
      }
      if (transaction == null)
      {
        throw new ArgumentNullException(nameof(transaction));
      }
      if (metadata == null ||
          string.IsNullOrWhiteSpace(metadata.Name) ||
          metadata.SquareFeet < 0.0 ||
          !IsFinite(metadata.SquareFeet) ||
          !IsFinite(metadata.BasePoint.X) ||
          !IsFinite(metadata.BasePoint.Y) ||
          !IsFinite(metadata.BasePoint.Z) ||
          !IsFinite(metadata.RelativeLocation.X) ||
          !IsFinite(metadata.RelativeLocation.Y))
      {
        throw new ArgumentException(
          "Room metadata requires a name and valid numeric values.",
          nameof(metadata));
      }

      if (!polyline.IsWriteEnabled)
      {
        polyline.UpgradeOpen();
      }
      if (polyline.ExtensionDictionary.IsNull)
      {
        polyline.CreateExtensionDictionary();
      }

      DBDictionary dictionary = transaction.GetObject(
        polyline.ExtensionDictionary,
        OpenMode.ForWrite,
        false) as DBDictionary;
      if (dictionary == null)
      {
        throw new InvalidOperationException(
          "Unable to open the room polyline extension dictionary.");
      }

      ResultBuffer data = new ResultBuffer(
        new TypedValue((int)DxfCode.Int32, RecordVersion),
        new TypedValue((int)DxfCode.Text, metadata.Name.Trim()),
        new TypedValue((int)DxfCode.Real, metadata.SquareFeet),
        new TypedValue((int)DxfCode.Real, metadata.BasePoint.X),
        new TypedValue((int)DxfCode.Real, metadata.BasePoint.Y),
        new TypedValue((int)DxfCode.Real, metadata.BasePoint.Z),
        new TypedValue((int)DxfCode.Real, metadata.RelativeLocation.X),
        new TypedValue((int)DxfCode.Real, metadata.RelativeLocation.Y));

      if (dictionary.Contains(RecordKey))
      {
        Xrecord existing = transaction.GetObject(
          dictionary.GetAt(RecordKey),
          OpenMode.ForWrite,
          false) as Xrecord;
        if (existing == null)
        {
          throw new InvalidOperationException(
            "The existing room metadata record is invalid.");
        }
        existing.Data = data;
      }
      else
      {
        Xrecord record = new Xrecord { Data = data };
        dictionary.SetAt(RecordKey, record);
        transaction.AddNewlyCreatedDBObject(record, true);
      }
    }

    private static bool IsFinite(double value)
    {
      return !double.IsNaN(value) && !double.IsInfinity(value);
    }
  }
}
