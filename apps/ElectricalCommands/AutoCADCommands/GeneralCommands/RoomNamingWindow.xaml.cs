using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
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
    private readonly List<AreaLabelRoomItem> _rooms;
    private readonly string _drawingDirectory;
    private readonly Editor _editor;

    internal event Action<ObjectId> SelectedBoundaryChanged;

    internal RoomNamingWindow(
      List<AreaLabelRoomItem> rooms,
      string drawingDirectory = null,
      Editor editor = null)
    {
      _rooms = rooms ?? new List<AreaLabelRoomItem>();
      _drawingDirectory = drawingDirectory ?? string.Empty;
      _editor = editor;
      InitializeComponent();
      RoomsDataGrid.ItemsSource = _rooms;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      if (_rooms.Count > 0)
      {
        RoomsDataGrid.SelectedIndex = 0;
        RoomsDataGrid.ScrollIntoView(_rooms[0]);
      }
    }

    private void RoomsDataGrid_SelectionChanged(
      object sender,
      SelectionChangedEventArgs e)
    {
      if (!(RoomsDataGrid.SelectedItem is AreaLabelRoomItem selected))
      {
        SelectedRoomTextBlock.Text = string.Empty;
        return;
      }

      SelectedRoomTextBlock.Text =
        $"{selected.SourceLabel}  |  {selected.SquareFootageText}  |  " +
        $"AutoCAD handle {selected.SourceHandle}";
      StatusTextBlock.Text =
        "The selected boundary is highlighted in the drawing. " +
        "Its saved name is preloaded when metadata exists; double-click " +
        "the Room Name cell to edit it.";
      StatusTextBlock.Foreground = Brushes.SlateGray;
      SelectedBoundaryChanged?.Invoke(selected.ObjectId);
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

    private void Save_Click(object sender, RoutedEventArgs e)
    {
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

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
