using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
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

  public class AreaLabelGroupExportItem
  {
    [JsonProperty("Group")]
    public string Group { get; set; }

    [JsonProperty("Polylines")]
    public int Polylines { get; set; }

    [JsonProperty("SquareFeet")]
    public double SquareFeet { get; set; }

    [JsonProperty("PerimeterFeet")]
    public double PerimeterFeet { get; set; }
  }

  internal sealed class AreaLabelRoomItem : INotifyPropertyChanged
  {
    private string _roomName = string.Empty;
    private string _group = string.Empty;

    internal ObjectId ObjectId { get; set; }
    internal string DefaultRoomName { get; set; } = string.Empty;
    internal double SquareFeet { get; set; }
    internal double PerimeterFeet { get; set; }

    public string SourceLabel { get; set; } = string.Empty;
    public string SourceHandle { get; set; } = string.Empty;
    public string SquareFootageText =>
      SquareFeet.ToString("0.##", CultureInfo.CurrentCulture) + " sq ft";
    public string PerimeterText =>
      PerimeterFeet.ToString("0.##", CultureInfo.CurrentCulture) + " ft";

    // Area group such as INTERIOR or EXTERIOR, totaled separately for T24 forms.
    public string Group
    {
      get => _group;
      set
      {
        if (string.Equals(_group, value, StringComparison.Ordinal))
        {
          return;
        }
        _group = value ?? string.Empty;
        OnPropertyChanged();
      }
    }

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

  internal sealed class AreaLabelGroupTotal
  {
    internal const string UngroupedName = "(UNGROUPED)";
    internal const string TotalName = "TOTAL";

    internal List<AreaLabelRoomItem> Rooms { get; } = new List<AreaLabelRoomItem>();
    internal double SquareFeet => Rooms.Sum(r => r.SquareFeet);
    internal double PerimeterFeet => Rooms.Sum(r => r.PerimeterFeet);

    public string GroupName { get; set; } = string.Empty;
    public bool IsTotal { get; set; }
    public int Count => Rooms.Count;
    public string AreaText =>
      Math.Round(SquareFeet, 2).ToString("0.##", CultureInfo.CurrentCulture) + " sq ft";
    public string PerimeterText =>
      Math.Round(PerimeterFeet, 2).ToString("0.##", CultureInfo.CurrentCulture) + " ft";

    internal static string NormalizeGroupName(string groupName)
    {
      return Regex.Replace(
        groupName ?? string.Empty,
        @"\s+",
        " ").Trim().ToUpperInvariant();
    }

    /// <summary>
    /// One total per named group (alphabetical), then ungrouped polylines,
    /// then the overall total.
    /// </summary>
    internal static List<AreaLabelGroupTotal> Build(IEnumerable<AreaLabelRoomItem> rooms)
    {
      List<AreaLabelRoomItem> roomList = rooms?.ToList() ?? new List<AreaLabelRoomItem>();
      List<AreaLabelGroupTotal> totals = roomList
        .GroupBy(r => NormalizeGroupName(r.Group), StringComparer.OrdinalIgnoreCase)
        .OrderBy(g => g.Key.Length == 0 ? 1 : 0)
        .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
        .Select(g =>
        {
          var total = new AreaLabelGroupTotal
          {
            GroupName = g.Key.Length == 0 ? UngroupedName : g.Key,
          };
          total.Rooms.AddRange(g);
          return total;
        })
        .ToList();

      var overall = new AreaLabelGroupTotal { GroupName = TotalName, IsTotal = true };
      overall.Rooms.AddRange(roomList);
      totals.Add(overall);
      return totals;
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

    private readonly ObservableCollection<AreaLabelGroupTotal> _groupTotals =
      new ObservableCollection<AreaLabelGroupTotal>();
    private bool _syncingGroupSelection;

    internal ObservableCollection<AreaLabelRoomItem> Rooms => _rooms;
    internal HashSet<ObjectId> RemovedObjectIds { get; } = new HashSet<ObjectId>();

    // Suggestions for the group pickers; groups already on the polylines are added.
    public ObservableCollection<string> GroupOptions { get; } =
      new ObservableCollection<string> { "INTERIOR", "EXTERIOR" };

    // Common base point for the saved room locations; null until one is
    // saved in the drawing or picked here, in which case it is prompted after review.
    internal Point3d? BasePoint { get; private set; }

    internal RoomNamingWindow(
      IEnumerable<AreaLabelRoomItem> rooms,
      string drawingDirectory = null,
      Editor editor = null,
      Database database = null,
      Point3d? basePoint = null)
    {
      _rooms = rooms != null
        ? new ObservableCollection<AreaLabelRoomItem>(rooms)
        : new ObservableCollection<AreaLabelRoomItem>();
      _drawingDirectory = drawingDirectory ?? string.Empty;
      _editor = editor;
      _database = database ?? editor?.Document?.Database;
      BasePoint = basePoint;
      InitializeComponent();
      UpdateBasePointText();
      RoomsDataGrid.ItemsSource = _rooms;
      GroupTotalsDataGrid.ItemsSource = _groupTotals;
      GroupComboBox.ItemsSource = GroupOptions;

      foreach (AreaLabelRoomItem room in _rooms)
      {
        room.PropertyChanged += Room_PropertyChanged;
      }
      _rooms.CollectionChanged += Rooms_CollectionChanged;

      SyncGroupOptions();
      UpdateRoomSummary();
    }

    private void Rooms_CollectionChanged(
      object sender,
      System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
      if (e.OldItems != null)
      {
        foreach (AreaLabelRoomItem room in e.OldItems)
        {
          room.PropertyChanged -= Room_PropertyChanged;
        }
      }
      if (e.NewItems != null)
      {
        foreach (AreaLabelRoomItem room in e.NewItems)
        {
          room.PropertyChanged += Room_PropertyChanged;
        }
      }
      SyncGroupOptions();
      UpdateGroupTotals();
    }

    private void Room_PropertyChanged(object sender, PropertyChangedEventArgs e)
    {
      if (e.PropertyName == nameof(AreaLabelRoomItem.Group))
      {
        UpdateGroupTotals();
      }
    }

    private void SyncGroupOptions()
    {
      foreach (string group in _rooms
        .Select(r => AreaLabelGroupTotal.NormalizeGroupName(r.Group))
        .Where(g => g.Length > 0)
        .Distinct(StringComparer.OrdinalIgnoreCase))
      {
        if (!GroupOptions.Contains(group, StringComparer.OrdinalIgnoreCase))
        {
          GroupOptions.Add(group);
        }
      }
    }

    private void UpdateGroupTotals()
    {
      if (GroupTotalsDataGrid == null) return;
      _syncingGroupSelection = true;
      try
      {
        _groupTotals.Clear();
        foreach (AreaLabelGroupTotal total in AreaLabelGroupTotal.Build(_rooms))
        {
          _groupTotals.Add(total);
        }
      }
      finally
      {
        _syncingGroupSelection = false;
      }
    }

    private void GroupTotalsDataGrid_SelectionChanged(
      object sender,
      SelectionChangedEventArgs e)
    {
      if (_syncingGroupSelection ||
          !(GroupTotalsDataGrid.SelectedItem is AreaLabelGroupTotal total))
      {
        return;
      }

      // Selecting a group total selects its polylines, which highlights them
      // in the drawing through the normal row selection.
      _syncingGroupSelection = true;
      try
      {
        RoomsDataGrid.SelectedItems.Clear();
        foreach (AreaLabelRoomItem room in total.Rooms)
        {
          RoomsDataGrid.SelectedItems.Add(room);
        }
        if (total.Rooms.Count > 0)
        {
          RoomsDataGrid.ScrollIntoView(total.Rooms[0]);
        }
      }
      finally
      {
        _syncingGroupSelection = false;
      }
    }

    private void AssignGroup_Click(object sender, RoutedEventArgs e)
    {
      string group = AreaLabelGroupTotal.NormalizeGroupName(GroupComboBox.Text);
      if (group.Length == 0)
      {
        StatusTextBlock.Text = "Type or pick a group name (for example INTERIOR or EXTERIOR) first.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }
      SetGroupOnSelected(group);
    }

    private void ClearGroup_Click(object sender, RoutedEventArgs e)
    {
      SetGroupOnSelected(string.Empty);
    }

    private void SetGroupOnSelected(string group)
    {
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

      var selectedList = RoomsDataGrid.SelectedItems.Cast<AreaLabelRoomItem>().ToList();
      if (selectedList.Count == 0)
      {
        StatusTextBlock.Text = "Select one or more room rows to group.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      foreach (AreaLabelRoomItem room in selectedList)
      {
        room.Group = group;
      }
      SyncGroupOptions();

      StatusTextBlock.Text = group.Length == 0
        ? $"Cleared the group on {selectedList.Count} polyline(s)."
        : $"Assigned {selectedList.Count} polyline(s) to {group}.";
      StatusTextBlock.Foreground = Brushes.SeaGreen;
    }

    private void RoomsDataGrid_CellEditEnding(
      object sender,
      DataGridCellEditEndingEventArgs e)
    {
      // The edited value is pushed to the row after this event, so pick up a
      // newly typed group name once the edit has finished.
      Dispatcher.BeginInvoke(new Action(SyncGroupOptions));
    }

    private void CopyTotals_Click(object sender, RoutedEventArgs e)
    {
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
      RoomsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

      var lines = new List<string> { "Group\tPolylines\tArea (sq ft)\tPerimeter (ft)" };
      foreach (AreaLabelGroupTotal total in AreaLabelGroupTotal.Build(_rooms))
      {
        lines.Add(string.Join(
          "\t",
          total.GroupName,
          total.Count.ToString(CultureInfo.InvariantCulture),
          Math.Round(total.SquareFeet, 2).ToString("0.##", CultureInfo.InvariantCulture),
          Math.Round(total.PerimeterFeet, 2).ToString("0.##", CultureInfo.InvariantCulture)));
      }

      try
      {
        Clipboard.SetText(string.Join(Environment.NewLine, lines));
        StatusTextBlock.Text = "Copied the group totals to the clipboard (paste into Excel or the T24 form).";
        StatusTextBlock.Foreground = Brushes.SeaGreen;
      }
      catch (System.Exception ex)
      {
        StatusTextBlock.Text = $"Copy failed: {ex.Message}";
        StatusTextBlock.Foreground = Brushes.Firebrick;
      }
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
      double totalPerimeter = Math.Round(_rooms.Sum(r => r.PerimeterFeet), 2);
      RoomSummaryTextBlock.Text =
        $"{_rooms.Count} Room Boundaries ({totalSqFt:0.##} sq ft, {totalPerimeter:0.##} ft perimeter total)";
      UpdateGroupTotals();
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
        string groupText = string.IsNullOrWhiteSpace(selected.Group)
          ? "no group"
          : AreaLabelGroupTotal.NormalizeGroupName(selected.Group);
        SelectedRoomTextBlock.Text =
          $"{selected.SourceLabel}  |  {selected.SquareFootageText}  |  " +
          $"{selected.PerimeterText} perimeter  |  {groupText}  |  " +
          $"AutoCAD handle {selected.SourceHandle}";
      }
      else
      {
        double sumSqFt = Math.Round(selectedList.Sum(r => r.SquareFeet), 2);
        double sumPerimeter = Math.Round(selectedList.Sum(r => r.PerimeterFeet), 2);
        SelectedRoomTextBlock.Text =
          $"{selectedList.Count} rooms selected  |  {sumSqFt:0.##} sq ft  |  " +
          $"{sumPerimeter:0.##} ft perimeter total";
      }

      StatusTextBlock.Text =
        "The selected boundary(ies) are highlighted in the drawing. " +
        "Double-click a Room Name cell to edit it.";
      StatusTextBlock.Foreground = Brushes.SlateGray;

      if (!_syncingGroupSelection)
      {
        GroupTotalsDataGrid.UnselectAll();
      }

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

        // Group totals go in their own file so AreaLabel.json keeps the
        // room-type format SUMTEXT also writes.
        List<AreaLabelGroupExportItem> groupExportList = AreaLabelGroupTotal.Build(_rooms)
          .Select(total => new AreaLabelGroupExportItem
          {
            Group = total.GroupName,
            Polylines = total.Count,
            SquareFeet = Math.Round(total.SquareFeet, 2),
            PerimeterFeet = Math.Round(total.PerimeterFeet, 2),
          })
          .ToList();
        string groupsPath = Path.Combine(
          Path.GetDirectoryName(targetPath) ?? string.Empty,
          "AreaLabelGroups.json");
        File.WriteAllText(
          groupsPath,
          JsonConvert.SerializeObject(groupExportList, Formatting.Indented));

        StatusTextBlock.Text =
          $"Exported {exportList.Count - 1} room type(s) ({totalSquareFeet:0.##} sq ft total) to {Path.GetFileName(targetPath)} " +
          $"and {groupExportList.Count - 1} group total(s) to {Path.GetFileName(groupsPath)}.";
        StatusTextBlock.Foreground = Brushes.SeaGreen;

        _editor?.WriteMessage($"\nExported room information to: {targetPath}");
        _editor?.WriteMessage($"\nExported group totals to: {groupsPath}");
      }
      catch (System.Exception ex)
      {
        StatusTextBlock.Text = $"Export failed: {ex.Message}";
        StatusTextBlock.Foreground = Brushes.Firebrick;
      }
    }

    private void UpdateBasePointText()
    {
      if (BasePoint.HasValue)
      {
        BasePointTextBlock.Text =
          $"Base point: {GeneralCommands.FormatAreaLabelPoint(BasePoint.Value)}. " +
          "No drawing text will be created.";
        BasePointButton.Content = "Change Base Point";
      }
      else
      {
        BasePointTextBlock.Text =
          "The base point will be requested after this review unless you set it here. " +
          "No drawing text will be created.";
        BasePointButton.Content = "Set Base Point";
      }
    }

    private void BasePoint_Click(object sender, RoutedEventArgs e)
    {
      if (_editor == null)
      {
        StatusTextBlock.Text = "AutoCAD editor is not available to pick a base point.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      EditorUserInteraction interaction = null;
      try
      {
        interaction = _editor.StartUserInteraction(this);

        string prompt = BasePoint.HasValue
          ? $"\nSpecify new common base point <{GeneralCommands.FormatAreaLabelPoint(BasePoint.Value)}>: "
          : "\nSpecify the common base point for the saved room boundaries: ";
        PromptPointResult result = _editor.GetPoint(new PromptPointOptions(prompt));
        if (result.Status != PromptStatus.OK)
        {
          return;
        }

        BasePoint = result.Value.TransformBy(_editor.CurrentUserCoordinateSystem);
        UpdateBasePointText();
        StatusTextBlock.Text =
          $"Base point set to {GeneralCommands.FormatAreaLabelPoint(BasePoint.Value)}. " +
          "Room locations are saved relative to it when you save.";
        StatusTextBlock.Foreground = Brushes.SeaGreen;
      }
      catch (System.Exception ex)
      {
        StatusTextBlock.Text = $"Failed to set the base point: {ex.Message}";
        StatusTextBlock.Foreground = Brushes.Firebrick;
      }
      finally
      {
        interaction?.End();
        Focus();
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
            string group = string.Empty;

            if (RoomBoundaryMetadataStore.TryRead(
              polyline,
              transaction,
              out var existingMetadata))
            {
              roomName = existingMetadata.Name;
              group = existingMetadata.Group;
            }

            var newItem = new AreaLabelRoomItem
            {
              ObjectId = objectId,
              SourceLabel = defaultName,
              SourceHandle = handleStr,
              DefaultRoomName = defaultName,
              RoomName = roomName,
              SquareFeet = squareFeet,
              PerimeterFeet = GeneralCommands.CalculateAreaLabelPerimeterFeet(polyline),
              Group = group,
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

    private void ZoomToSelected_Click(object sender, RoutedEventArgs e)
    {
      var selectedList = RoomsDataGrid.SelectedItems.Cast<AreaLabelRoomItem>().ToList();
      if (selectedList.Count == 0)
      {
        StatusTextBlock.Text = "Please select one or more room rows to zoom to.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      Database db = _database ?? _editor?.Document?.Database;
      if (_editor == null || db == null)
      {
        StatusTextBlock.Text = "AutoCAD editor is not available to zoom.";
        StatusTextBlock.Foreground = Brushes.Firebrick;
        return;
      }

      try
      {
        Extents3d? combined = null;
        using (Transaction transaction = db.TransactionManager.StartOpenCloseTransaction())
        {
          foreach (AreaLabelRoomItem room in selectedList)
          {
            if (room.ObjectId.IsNull || !room.ObjectId.IsValid || room.ObjectId.IsErased)
            {
              continue;
            }

            if (!(transaction.GetObject(room.ObjectId, OpenMode.ForRead, false) is Entity entity))
            {
              continue;
            }

            Extents3d extents = entity.GeometricExtents;
            if (combined.HasValue)
            {
              Extents3d merged = combined.Value;
              merged.AddExtents(extents);
              combined = merged;
            }
            else
            {
              combined = extents;
            }
          }
        }

        if (!combined.HasValue)
        {
          StatusTextBlock.Text = "The selected polyline(s) no longer exist in the drawing.";
          StatusTextBlock.Foreground = Brushes.Firebrick;
          return;
        }

        ZoomEditorToExtents(combined.Value);

        StatusTextBlock.Text = $"Zoomed to {selectedList.Count} selected room polyline(s).";
        StatusTextBlock.Foreground = Brushes.SlateGray;
      }
      catch (System.Exception ex)
      {
        StatusTextBlock.Text = $"Zoom failed: {ex.Message}";
        StatusTextBlock.Foreground = Brushes.Firebrick;
      }
    }

    private void ZoomEditorToExtents(Extents3d extents)
    {
      using (ViewTableRecord view = _editor.GetCurrentView())
      {
        Matrix3d worldToDisplay = Matrix3d.PlaneToWorld(view.ViewDirection)
          .PreMultiplyBy(Matrix3d.Displacement(view.Target - Point3d.Origin))
          .PreMultiplyBy(Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target))
          .Inverse();
        extents.TransformBy(worldToDisplay);

        // Leave a margin around the boundary so its edges are not at the screen border.
        double width = (extents.MaxPoint.X - extents.MinPoint.X) * 1.2;
        double height = (extents.MaxPoint.Y - extents.MinPoint.Y) * 1.2;
        double viewRatio = view.Width / view.Height;
        if (width > height * viewRatio)
        {
          height = width / viewRatio;
        }
        else
        {
          width = height * viewRatio;
        }

        view.CenterPoint = new Point2d(
          (extents.MinPoint.X + extents.MaxPoint.X) / 2.0,
          (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0);
        view.Width = width;
        view.Height = height;
        _editor.SetCurrentView(view);
      }
      Autodesk.AutoCAD.ApplicationServices.Application.UpdateScreen();
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
        room.Group = AreaLabelGroupTotal.NormalizeGroupName(room.Group);
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
