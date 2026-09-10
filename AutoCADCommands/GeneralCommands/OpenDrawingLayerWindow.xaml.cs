using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ElectricalCommands
{
  internal enum OpenDrawingLayerAction
  {
    Untouched,
    Freeze,
    Thaw
  }

  internal sealed class OpenDrawingLayerRow : INotifyPropertyChanged
  {
    private OpenDrawingLayerAction _plannedAction;

    public OpenDrawingLayerRow(
      string name,
      int presentCount,
      int selectedDrawingCount,
      int frozenCount,
      int currentCount,
      OpenDrawingLayerAction plannedAction)
    {
      Name = name;
      PresenceSummary = $"{presentCount} of {selectedDrawingCount}";
      StateSummary = frozenCount == 0
        ? "Thawed in all"
        : frozenCount == presentCount
          ? "Frozen in all"
          : $"Mixed: {frozenCount} frozen";
      CurrentSummary = currentCount == 0
        ? "No"
        : currentCount == 1
          ? "Yes, in 1"
          : $"Yes, in {currentCount}";
      _plannedAction = plannedAction;
    }

    public string Name { get; }
    public string PresenceSummary { get; }
    public string StateSummary { get; }
    public string CurrentSummary { get; }

    public OpenDrawingLayerAction PlannedAction
    {
      get => _plannedAction;
      set
      {
        if (_plannedAction == value) return;
        _plannedAction = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PlannedAction)));
      }
    }

    public event PropertyChangedEventHandler PropertyChanged;
  }

  public partial class OpenDrawingLayerWindow : Window
  {
    private readonly IReadOnlyList<OpenDrawingLayerSnapshot> _drawings;
    private readonly Dictionary<string, OpenDrawingLayerAction> _actionsByLayer =
      new Dictionary<string, OpenDrawingLayerAction>(StringComparer.OrdinalIgnoreCase);
    private ObservableCollection<OpenDrawingLayerRow> _rows =
      new ObservableCollection<OpenDrawingLayerRow>();
    private ICollectionView _layerView;

    internal OpenDrawingLayerWindow(IReadOnlyList<OpenDrawingLayerSnapshot> drawings)
    {
      InitializeComponent();
      _drawings = drawings ?? throw new ArgumentNullException(nameof(drawings));
      DrawingListBox.ItemsSource = _drawings;
      RebuildLayerRows();
    }

    internal IReadOnlyList<OpenDrawingLayerSnapshot> SelectedDrawings =>
      _drawings.Where(drawing => drawing.IsSelected).ToList();

    internal IReadOnlyDictionary<string, OpenDrawingLayerAction> LayerActions =>
      _rows
        .Where(row => row.PlannedAction != OpenDrawingLayerAction.Untouched)
        .ToDictionary(row => row.Name, row => row.PlannedAction, StringComparer.OrdinalIgnoreCase);

    internal bool SaveDrawingsAfterApplying => SaveDrawingsCheckBox.IsChecked == true;

    private void RebuildLayerRows()
    {
      List<OpenDrawingLayerSnapshot> selectedDrawings = _drawings
        .Where(drawing => drawing.IsSelected)
        .ToList();

      var layerNames = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
      foreach (OpenDrawingLayerSnapshot drawing in selectedDrawings)
      {
        foreach (string layerName in drawing.Layers.Keys)
        {
          layerNames.Add(layerName);
        }
      }

      _rows = new ObservableCollection<OpenDrawingLayerRow>(
        layerNames.Select(layerName => CreateLayerRow(layerName, selectedDrawings)));

      _layerView = CollectionViewSource.GetDefaultView(_rows);
      _layerView.Filter = MatchesLayerFilter;
      LayerGrid.ItemsSource = _layerView;
      UpdateStatus();
    }

    private OpenDrawingLayerRow CreateLayerRow(
      string layerName,
      IReadOnlyCollection<OpenDrawingLayerSnapshot> selectedDrawings)
    {
      List<OpenDrawingLayerState> states = selectedDrawings
        .Where(drawing => drawing.Layers.ContainsKey(layerName))
        .Select(drawing => drawing.Layers[layerName])
        .ToList();

      _actionsByLayer.TryGetValue(layerName, out OpenDrawingLayerAction plannedAction);
      return new OpenDrawingLayerRow(
        layerName,
        states.Count,
        selectedDrawings.Count,
        states.Count(state => state.IsFrozen),
        states.Count(state => state.IsCurrent),
        plannedAction);
    }

    private bool MatchesLayerFilter(object item)
    {
      var row = item as OpenDrawingLayerRow;
      if (row == null) return false;

      string query = SearchTextBox?.Text?.Trim();
      return string.IsNullOrWhiteSpace(query) ||
             row.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private void SetActionForSelectedLayers(OpenDrawingLayerAction action)
    {
      List<OpenDrawingLayerRow> selectedRows = LayerGrid.SelectedItems
        .OfType<OpenDrawingLayerRow>()
        .ToList();

      if (selectedRows.Count == 0)
      {
        MessageBox.Show(
          "Select one or more layer rows first.",
          "Open Drawing Layer Manager",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
        return;
      }

      foreach (OpenDrawingLayerRow row in selectedRows)
      {
        row.PlannedAction = action;
        if (action == OpenDrawingLayerAction.Untouched)
        {
          _actionsByLayer.Remove(row.Name);
        }
        else
        {
          _actionsByLayer[row.Name] = action;
        }
      }

      UpdateStatus();
    }

    private void UpdateStatus()
    {
      int selectedDrawingCount = _drawings.Count(drawing => drawing.IsSelected);
      int plannedCount = _rows.Count(row => row.PlannedAction != OpenDrawingLayerAction.Untouched);
      StatusTextBlock.Text =
        $"{selectedDrawingCount} drawing(s) · {_rows.Count} layer(s) · {plannedCount} planned";
      ApplyButton.IsEnabled = selectedDrawingCount > 0 && plannedCount > 0;
    }

    private void DrawingSelectionChanged(object sender, RoutedEventArgs e)
    {
      if (sender is CheckBox checkBox && checkBox.DataContext is OpenDrawingLayerSnapshot drawing)
      {
        drawing.IsSelected = checkBox.IsChecked == true;
      }

      RebuildLayerRows();
    }

    private void SelectAllDrawings_Click(object sender, RoutedEventArgs e)
    {
      foreach (OpenDrawingLayerSnapshot drawing in _drawings)
      {
        drawing.IsSelected = true;
      }

      DrawingListBox.Items.Refresh();
      RebuildLayerRows();
    }

    private void SelectNoDrawings_Click(object sender, RoutedEventArgs e)
    {
      foreach (OpenDrawingLayerSnapshot drawing in _drawings)
      {
        drawing.IsSelected = false;
      }

      DrawingListBox.Items.Refresh();
      RebuildLayerRows();
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
      _layerView?.Refresh();
    }

    private void FreezeSelected_Click(object sender, RoutedEventArgs e)
    {
      SetActionForSelectedLayers(OpenDrawingLayerAction.Freeze);
    }

    private void ThawSelected_Click(object sender, RoutedEventArgs e)
    {
      SetActionForSelectedLayers(OpenDrawingLayerAction.Thaw);
    }

    private void LeaveSelectedUntouched_Click(object sender, RoutedEventArgs e)
    {
      SetActionForSelectedLayers(OpenDrawingLayerAction.Untouched);
    }

    private void ClearAllActions_Click(object sender, RoutedEventArgs e)
    {
      _actionsByLayer.Clear();
      foreach (OpenDrawingLayerRow row in _rows)
      {
        row.PlannedAction = OpenDrawingLayerAction.Untouched;
      }

      UpdateStatus();
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
      if (SelectedDrawings.Count == 0)
      {
        MessageBox.Show(
          "Select at least one open drawing.",
          "Open Drawing Layer Manager",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
        return;
      }

      if (LayerActions.Count == 0)
      {
        MessageBox.Show(
          "Assign Freeze or Thaw to at least one layer.",
          "Open Drawing Layer Manager",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
        return;
      }

      DialogResult = true;
      Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
      DialogResult = false;
      Close();
    }
  }
}
