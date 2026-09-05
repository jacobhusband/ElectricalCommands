using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ElectricalCommands
{
  public class AreaLabelExportItem
  {
    [JsonProperty("RoomType")]
    public string RoomType { get; set; }

    [JsonProperty("SquareFeet")]
    public double SquareFeet { get; set; }
  }

  internal sealed class AreaLabelRoomItem : INotifyPropertyChanged
  {
    private string _roomName = string.Empty;

    internal ObjectId ObjectId { get; set; }
    internal string DefaultRoomName { get; set; } = string.Empty;
    internal double SquareFeet { get; set; }

    public string SourceLabel { get; set; } = string.Empty;
    public string SourceHandle { get; set; } = string.Empty;
    public string SquareFootageText =>
      SquareFeet.ToString("0.##", CultureInfo.CurrentCulture) + " sq ft";

    public string RoomName
    {
      get => _roomName;
      set
      {
        if (string.Equals(_roomName, value, StringComparison.Ordinal))
        {
          return;
        }
        _roomName = value ?? string.Empty;
        OnPropertyChanged();
      }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
      PropertyChanged?.Invoke(
        this,
        new PropertyChangedEventArgs(propertyName));
    }
  }

  public partial class RoomNamingWindow : Window
  {
    private readonly ObservableCollection<AreaLabelRoomItem> _rooms;
    private readonly string _drawingDirectory;
    private readonly Editor _editor;
    private readonly Database _database;

    internal event Action<ObjectId> SelectedBoundaryChanged;
    internal event Action<List<ObjectId>> SelectedBoundariesChanged;

    internal ObservableCollection<AreaLabelRoomItem> Rooms => _rooms;
    internal HashSet<ObjectId> RemovedObjectIds { get; } = new HashSet<ObjectId>();

    internal RoomNamingWindow(
      IEnumerable<AreaLabelRoomItem> rooms,
      string drawingDirectory = null,
      Editor editor = null,
      Database database = null)
    {
      _rooms = rooms != null
        ? new ObservableCollection<AreaLabelRoomItem>(rooms)
        : new ObservableCollection<AreaLabelRoomItem>();
      _drawingDirectory = drawingDirectory ?? string.Empty;
      _editor = editor;
      _database = database ?? editor?.Document?.Database;
      InitializeComponent();
      RoomsDataGrid.ItemsSource = _rooms;
      UpdateRoomSummary();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      UpdateRoomSummary();
      if (_rooms.Count > 0)
      {
        RoomsDataGrid.SelectedIndex = 0;
        RoomsDataGrid.ScrollIntoView(_rooms[0]);
      }
    }

    private void UpdateRoomSummary()
    {
      if (RoomSummaryTextBlock == null) return;
      double totalSqFt = Math.Round(_rooms.Sum(r => r.SquareFeet), 2);
      RoomSummaryTextBlock.Text = $"{_rooms.Count} Room Boundaries ({totalSqFt:0.##} sq ft total)";
    }

    private void RoomsDataGrid_SelectionChanged(
      object sender,
      SelectionChangedEventArgs e)
    {
      var selectedList = RoomsDataGrid.SelectedItems.Cast<AreaLabelRoomItem>().ToList();
      if (selectedList.Count == 0)
      {
        SelectedRoomTextBlock.Text = string.Empty;
        SelectedBoundaryChanged?.Invoke(ObjectId.Null);
        SelectedBoundariesChanged?.Invoke(new List<ObjectId>());
        return;
      }

      if (selectedList.Count == 1)
      {
        var selected = selectedList[0];
        SelectedRoomTextBlock.Text =
          $"{selected.SourceLabel}  |  {selected.SquareFootageText}  |  " +
          $"AutoCAD handle {selected.SourceHandle}";
      }
      else
      {
        double sumSqFt = Math.Round(selectedList.Sum(r => r.SquareFeet), 2);
        SelectedRoomTextBlock.Text =
          $"{selectedList.Count} rooms selected  |  {sumSqFt:0.##} sq ft total";
      }

      StatusTextBlock.Text =
        "The selected boundary(ies) are highlighted in the drawing. " +
        "Double-click a Room Name cell to edit it.";
      StatusTextBlock.Foreground = Brushes.SlateGray;

      SelectedBoundaryChanged?.Invoke(selectedList[0].ObjectId);
      SelectedBoundariesChanged?.Invoke(selectedList.Select(r => r.ObjectId).ToList());
    }

    private void ExportJson_Click(object sender, RoutedEventArgs e)
    {
      try
      {
        RoomsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        RoomsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        if (_rooms == null || _rooms.Count == 0)
        {
          StatusTextBlock.Text = "No rooms available to export.";
          StatusTextBlock.Foreground = Brushes.Firebrick;
          return;
        }

        List<AreaLabelExportItem> exportList = _rooms
          .GroupBy(
            r => string.IsNullOrWhiteSpace(r.RoomName)
              ? NormalizeRoomName(r.DefaultRoomName)
              : NormalizeRoomName(r.RoomName),
            StringComparer.OrdinalIgnoreCase)
          .Select(group => new AreaLabelExportItem
          {
            RoomType = group.Key,
            SquareFeet = Math.Round(group.Sum(r => r.SquareFeet), 2)
          })
          .ToList();

        double totalSquareFeet = Math.Round(_rooms.Sum(r => r.SquareFeet), 2);
        exportList.Add(new AreaLabelExportItem
        {
          RoomType = "TOTAL",
          SquareFeet = totalSquareFeet
        });

        string targetPath = null;
        if (!string.IsNullOrWhiteSpace(_drawingDirectory) && Directory.Exists(_drawingDirectory))
        {
          targetPath = Path.Combine(_drawingDirectory, "AreaLabel.json");
        }
        else
        {
          var saveDialog = new Microsoft.Win32.SaveFileDialog
          {
            FileName = "AreaLabel.json",
            DefaultExt = ".json",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            Title = "Save AreaLabel JSON"
          };
          if (saveDialog.ShowDialog(this) == true)
          {
            targetPath = saveDialog.FileName;
          }
        }

        if (string.IsNullOrWhiteSpace(targetPath))
        {
          StatusTextBlock.Text = "Export canceled (no valid destination selected).";
          StatusTextBlock.Foreground = Brushes.SlateGray;
          return;
        }

        string json = JsonConvert.SerializeObject(exportList, Formatting.Indented);
        File.WriteAllText(targetPath, json);

        StatusTextBlock.Text =
          $"Exported {exportList.Count - 1} room type(s) ({totalSquareFeet:0.##} sq ft total) to {Path.GetFileName(targetPath)}.";
        StatusTextBlock.Foreground = Brushes.SeaGreen;

        _editor?.WriteMessage($"\nExported room information to: {targetPath}");
      }
      catch (System.Exception ex)
      {
        StatusTextBlock.Text = $"Export failed: {ex.Message}";
        StatusTextBlock.Foreground = Brushes.Firebrick;
      }
    }

    private void AddPolylines_Click(object sender, RoutedEventArgs e)
    {
      if (_editor == null)
      {
        StatusTextBlock.Text = "AutoCAD editor is not available to add polylines.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      Database db = _database ?? _editor.Document?.Database;
      if (db == null)
      {
        StatusTextBlock.Text = "AutoCAD database is not available.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      EditorUserInteraction interaction = null;
      try
      {
        interaction = _editor.StartUserInteraction(this);

        PromptSelectionOptions options = new PromptSelectionOptions
        {
          MessageForAdding = "\nSelect additional room polylines: ",
          AllowDuplicates = false,
          RejectObjectsOnLockedLayers = true,
        };
        SelectionFilter filter = new SelectionFilter(
          new[]
          {
            new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
          });

        PromptSelectionResult selectionResult = _editor.GetSelection(options, filter);
        if (selectionResult.Status != PromptStatus.OK ||
            selectionResult.Value == null ||
            selectionResult.Value.Count == 0)
        {
          return;
        }

        HashSet<ObjectId> existingIds = new HashSet<ObjectId>(_rooms.Select(r => r.ObjectId));
        HashSet<string> existingHandles = new HashSet<string>(
          _rooms.Select(r => r.SourceHandle),
          StringComparer.OrdinalIgnoreCase);

        int addedCount = 0;
        AreaLabelRoomItem lastAdded = null;

        using (Transaction transaction = db.TransactionManager.StartOpenCloseTransaction())
        {
          foreach (ObjectId objectId in selectionResult.Value.GetObjectIds())
          {
            if (objectId.IsNull || !objectId.IsValid || objectId.IsErased)
            {
              continue;
            }

            if (existingIds.Contains(objectId))
            {
              continue;
            }

            Polyline polyline = transaction.GetObject(
              objectId,
              OpenMode.ForRead,
              false) as Polyline;
            if (polyline == null)
            {
              _editor.WriteMessage(
                $"\nIgnored object {objectId.Handle}; it is not a lightweight polyline.");
              continue;
            }
            if (polyline.NumberOfVertices < 3)
            {
              _editor.WriteMessage(
                $"\nIgnored polyline {polyline.Handle}; room boundaries must contain at least three vertices.");
              continue;
            }

            string handleStr = polyline.Handle.ToString();
            if (existingHandles.Contains(handleStr))
            {
              continue;
            }

            double squareFeet = polyline.Area / 144.0;
            int roomNumber = _rooms.Count + 1;
            string defaultName = $"Polyline {roomNumber}";
            string roomName = defaultName;

            if (RoomBoundaryMetadataStore.TryRead(
              polyline,
              transaction,
              out var existingMetadata))
            {
              roomName = existingMetadata.Name;
            }

            var newItem = new AreaLabelRoomItem
            {
              ObjectId = objectId,
              SourceLabel = defaultName,
              SourceHandle = handleStr,
              DefaultRoomName = defaultName,
              RoomName = roomName,
              SquareFeet = squareFeet,
            };

            RemovedObjectIds.Remove(objectId);

            _rooms.Add(newItem);
            existingIds.Add(objectId);
            existingHandles.Add(handleStr);
            lastAdded = newItem;
            addedCount++;
          }
        }

        if (addedCount > 0)
        {
          RenumberDefaultLabels();
          UpdateRoomSummary();
          if (lastAdded != null)
          {
            RoomsDataGrid.SelectedItem = lastAdded;
            RoomsDataGrid.ScrollIntoView(lastAdded);
          }
          StatusTextBlock.Text =
            $"Added {addedCount} new room polyline(s). Total rooms: {_rooms.Count}.";
          StatusTextBlock.Foreground = Brushes.SeaGreen;
        }
        else
        {
          StatusTextBlock.Text =
            "No new room polylines were added (selected polylines may already be in the list or invalid).";
          StatusTextBlock.Foreground = Brushes.SlateGray;
        }
      }
      catch (System.Exception ex)
      {
        StatusTextBlock.Text = $"Failed to add polylines: {ex.Message}";
        StatusTextBlock.Foreground = Brushes.Firebrick;
      }
      finally
      {
        interaction?.End();
        Focus();
      }
    }

    private void RemoveSelected_Click(object sender, RoutedEventArgs e)
    {
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

      var selectedList = RoomsDataGrid.SelectedItems.Cast<AreaLabelRoomItem>().ToList();
      if (selectedList.Count == 0)
      {
        StatusTextBlock.Text = "Please select one or more room rows to remove.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      foreach (var item in selectedList)
      {
        if (!item.ObjectId.IsNull)
        {
          RemovedObjectIds.Add(item.ObjectId);
        }
        _rooms.Remove(item);
      }

      SelectedBoundaryChanged?.Invoke(ObjectId.Null);
      SelectedBoundariesChanged?.Invoke(new List<ObjectId>());

      RenumberDefaultLabels();
      UpdateRoomSummary();

      if (_rooms.Count > 0)
      {
        int newIndex = Math.Min(RoomsDataGrid.SelectedIndex, _rooms.Count - 1);
        if (newIndex >= 0)
        {
          RoomsDataGrid.SelectedIndex = newIndex;
        }
      }
      else
      {
        SelectedRoomTextBlock.Text = string.Empty;
      }

      StatusTextBlock.Text =
        $"Removed {selectedList.Count} room polyline(s). Remaining: {_rooms.Count}.";
      StatusTextBlock.Foreground = Brushes.SlateGray;
    }

    private void RoomsDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
    {
      if (e.Key == Key.Delete)
      {
        if (e.OriginalSource is TextBox)
        {
          return;
        }

        RemoveSelected_Click(sender, e);
        e.Handled = true;
      }
    }

    private void RenumberDefaultLabels()
    {
      for (int i = 0; i < _rooms.Count; i++)
      {
        string label = $"Polyline {i + 1}";
        _rooms[i].SourceLabel = label;
        if (string.Equals(
          NormalizeRoomName(_rooms[i].RoomName),
          NormalizeRoomName(_rooms[i].DefaultRoomName),
          StringComparison.OrdinalIgnoreCase))
        {
          _rooms[i].DefaultRoomName = label;
          _rooms[i].RoomName = label;
        }
        else
        {
          _rooms[i].DefaultRoomName = label;
        }
      }
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

      if (_rooms.Count == 0)
      {
        StatusTextBlock.Text = "No room polylines to save. Add at least one polyline before saving.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      List<AreaLabelRoomItem> missingNames = _rooms
        .Where(room =>
          string.IsNullOrWhiteSpace(room.RoomName) ||
          string.Equals(
            NormalizeRoomName(room.RoomName),
            NormalizeRoomName(room.DefaultRoomName),
            StringComparison.OrdinalIgnoreCase))
        .ToList();
      if (missingNames.Count > 0)
      {
        AreaLabelRoomItem firstMissing = missingNames[0];
        RoomsDataGrid.SelectedItem = firstMissing;
        RoomsDataGrid.ScrollIntoView(firstMissing);
        StatusTextBlock.Text =
          $"Rename all room placeholders before saving. " +
          $"{missingNames.Count} room(s) still need names.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      foreach (AreaLabelRoomItem room in _rooms)
      {
        room.RoomName = NormalizeRoomName(room.RoomName);
      }

      DialogResult = true;
      Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
      DialogResult = false;
      Close();
    }

    private static string NormalizeRoomName(string roomName)
    {
      return Regex.Replace(
        roomName ?? string.Empty,
        @"\s+",
        " ").Trim().ToUpperInvariant();
    }
  }
}
