using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ElectricalCommands
{
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

    internal event Action<ObjectId> SelectedBoundaryChanged;

    internal RoomNamingWindow(List<AreaLabelRoomItem> rooms)
    {
      _rooms = rooms ?? new List<AreaLabelRoomItem>();
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
        "Double-click its Room Name cell to rename it.";
      StatusTextBlock.Foreground = Brushes.SlateGray;
      SelectedBoundaryChanged?.Invoke(selected.ObjectId);
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
