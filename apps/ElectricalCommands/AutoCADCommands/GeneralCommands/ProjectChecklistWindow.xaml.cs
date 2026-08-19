using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ElectricalCommands
{
  public partial class ProjectChecklistWindow : Window
  {
    private string _folderPath;
    private string _dwgPath;
    private List<ChecklistDefinition> _definitions;
    private FolderChecklistState _state;
    private string _selectedChecklistId = string.Empty;
    private bool _isUpdatingUi;

    internal ProjectChecklistWindow(
      string folderPath,
      string dwgPath,
      List<ChecklistDefinition> definitions,
      FolderChecklistState state
    )
    {
      InitializeComponent();

      _folderPath = folderPath ?? string.Empty;
      _dwgPath = dwgPath ?? string.Empty;
      _definitions = definitions ?? new List<ChecklistDefinition>();
      _state = state ?? new FolderChecklistState();

      _selectedChecklistId = !string.IsNullOrWhiteSpace(_state.LastActiveChecklistId)
        ? _state.LastActiveChecklistId
        : ProjectChecklistStore.LoadGlobalLastActiveChecklist();

      FolderPathTextBlock.Text = !string.IsNullOrWhiteSpace(_folderPath)
        ? _folderPath
        : "Unsaved / Unknown";
      DwgNameTextBlock.Text = !string.IsNullOrWhiteSpace(_dwgPath)
        ? Path.GetFileName(_dwgPath)
        : "Active Drawing";

      PopulateChecklists();
    }

    internal void SwitchFolderOrDrawing(string newFolderPath, string newDwgPath)
    {
      bool folderChanged = !string.Equals(_folderPath, newFolderPath, StringComparison.OrdinalIgnoreCase);
      _folderPath = newFolderPath ?? string.Empty;
      _dwgPath = newDwgPath ?? string.Empty;

      FolderPathTextBlock.Text = !string.IsNullOrWhiteSpace(_folderPath)
        ? _folderPath
        : "Unsaved / Unknown";
      DwgNameTextBlock.Text = !string.IsNullOrWhiteSpace(_dwgPath)
        ? Path.GetFileName(_dwgPath)
        : "Active Drawing";

      if (folderChanged)
      {
        _definitions = ProjectChecklistStore.LoadAllChecklists(_folderPath);
        _state = ProjectChecklistStore.LoadFolderState(_folderPath);
        _selectedChecklistId = !string.IsNullOrWhiteSpace(_state.LastActiveChecklistId)
          ? _state.LastActiveChecklistId
          : ProjectChecklistStore.LoadGlobalLastActiveChecklist();
      }

      PopulateChecklists();
      SetStatus($"Active drawing: {DwgNameTextBlock.Text}", isError: false);
    }

    private void PopulateChecklists()
    {
      _isUpdatingUi = true;
      try
      {
        ChecklistComboBox.Items.Clear();
        foreach (var def in _definitions)
        {
          ChecklistComboBox.Items.Add(new ChecklistComboItem(def));
        }

        if (ChecklistComboBox.Items.Count > 0)
        {
          int index = 0;
          string targetId = !string.IsNullOrWhiteSpace(_selectedChecklistId)
            ? _selectedChecklistId
            : (!string.IsNullOrWhiteSpace(_state.LastActiveChecklistId)
                ? _state.LastActiveChecklistId
                : ProjectChecklistStore.LoadGlobalLastActiveChecklist());

          if (!string.IsNullOrWhiteSpace(targetId))
          {
            for (int i = 0; i < _definitions.Count; i++)
            {
              if (string.Equals(_definitions[i].Id, targetId, StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(_definitions[i].Name, targetId, StringComparison.OrdinalIgnoreCase))
              {
                index = i;
                break;
              }
            }
          }
          ChecklistComboBox.SelectedIndex = index;
          _selectedChecklistId = _definitions[index].Id;
          _state.LastActiveChecklistId = _selectedChecklistId;
          ProjectChecklistStore.SaveGlobalLastActiveChecklist(_selectedChecklistId);
        }
      }
      finally
      {
        _isUpdatingUi = false;
      }

      RenderItems();
    }

    private ChecklistDefinition GetSelectedDefinition()
    {
      if (ChecklistComboBox.SelectedItem is ChecklistComboItem comboItem)
      {
        return comboItem.Definition;
      }
      return _definitions.FirstOrDefault(d => string.Equals(d.Id, _selectedChecklistId, StringComparison.OrdinalIgnoreCase))
        ?? _definitions.FirstOrDefault();
    }

    private FolderChecklistEntry GetOrCreateEntry(string checklistId)
    {
      if (string.IsNullOrWhiteSpace(checklistId))
      {
        return new FolderChecklistEntry();
      }

      if (!_state.Checklists.TryGetValue(checklistId, out var entry) || entry == null)
      {
        entry = new FolderChecklistEntry();
        _state.Checklists[checklistId] = entry;
      }
      return entry;
    }

    private void RenderItems()
    {
      var def = GetSelectedDefinition();
      if (def == null)
      {
        ItemsContainer.Children.Clear();
        ProgressLabel.Text = "No checklist selected.";
        ChecklistProgressBar.Value = 0;
        return;
      }

      _selectedChecklistId = def.Id;
      var entry = GetOrCreateEntry(def.Id);
      var completedSet = new HashSet<string>(entry.CompletedItems ?? new List<string>(), StringComparer.OrdinalIgnoreCase);

      string searchFilter = (SearchTextBox.Text ?? string.Empty).Trim().ToLowerInvariant();
      bool hideCompleted = HideCompletedCheckBox.IsChecked == true;

      // Calculate total actionable items and completed
      var actionableItems = def.Items.Where(item => !item.IsSubheader).ToList();
      int totalCount = actionableItems.Count;
      int completedCount = actionableItems.Count(item => completedSet.Contains(item.Id) || completedSet.Contains(item.Text));

      double percent = totalCount > 0 ? ((double)completedCount / totalCount) * 100.0 : 0.0;
      ChecklistProgressBar.Value = percent;
      ProgressLabel.Text = $"Progress: {completedCount} / {totalCount} completed ({percent:F0}%)";

      // Render UI
      ItemsContainer.Children.Clear();

      // Group items under current subheaders for visual structure
      string currentSubheader = string.Empty;
      var currentSectionItems = new List<ChecklistItem>();
      var sections = new List<Tuple<string, List<ChecklistItem>>>();

      foreach (var item in def.Items)
      {
        if (item.IsSubheader)
        {
          if (currentSectionItems.Count > 0 || !string.IsNullOrWhiteSpace(currentSubheader))
          {
            sections.Add(Tuple.Create(currentSubheader, currentSectionItems));
            currentSectionItems = new List<ChecklistItem>();
          }
          currentSubheader = item.Text;
        }
        else
        {
          currentSectionItems.Add(item);
        }
      }
      if (currentSectionItems.Count > 0 || !string.IsNullOrWhiteSpace(currentSubheader))
      {
        sections.Add(Tuple.Create(currentSubheader, currentSectionItems));
      }

      int visibleItemsCount = 0;

      foreach (var section in sections)
      {
        string sectionTitle = section.Item1;
        var sectionItems = section.Item2;

        // Filter items in section
        var matchingItems = sectionItems.Where(item =>
        {
          bool isCompleted = completedSet.Contains(item.Id) || completedSet.Contains(item.Text);
          if (hideCompleted && isCompleted)
          {
            return false;
          }
          if (!string.IsNullOrWhiteSpace(searchFilter))
          {
            return item.Text.ToLowerInvariant().Contains(searchFilter) ||
                   sectionTitle.ToLowerInvariant().Contains(searchFilter);
          }
          return true;
        }).ToList();

        if (matchingItems.Count == 0 && !string.IsNullOrWhiteSpace(searchFilter))
        {
          continue;
        }

        // Add section banner if title exists
        if (!string.IsNullOrWhiteSpace(sectionTitle))
        {
          int sectionTotal = sectionItems.Count;
          int sectionDone = sectionItems.Count(item => completedSet.Contains(item.Id) || completedSet.Contains(item.Text));
          ItemsContainer.Children.Add(CreateSectionHeader(sectionTitle, sectionDone, sectionTotal));
        }

        if (matchingItems.Count == 0 && hideCompleted)
        {
          continue;
        }

        foreach (var item in matchingItems)
        {
          visibleItemsCount++;
          bool isChecked = completedSet.Contains(item.Id) || completedSet.Contains(item.Text);
          ItemsContainer.Children.Add(CreateItemRow(def.Id, item, isChecked));
        }
      }

      if (visibleItemsCount == 0 && (hideCompleted || !string.IsNullOrWhiteSpace(searchFilter)))
      {
        var emptyText = new TextBlock
        {
          Text = hideCompleted && completedCount == totalCount && totalCount > 0
            ? "🎉 All items in this checklist are completed!"
            : "No items match your search/filter.",
          FontSize = 13,
          Foreground = new SolidColorBrush(Color.FromRgb(106, 115, 125)),
          HorizontalAlignment = HorizontalAlignment.Center,
          Margin = new Thickness(0, 30, 0, 30)
        };
        ItemsContainer.Children.Add(emptyText);
      }

      VisibleCountLabel.Text = visibleItemsCount < totalCount
        ? $"Showing {visibleItemsCount} of {totalCount} items"
        : string.Empty;
    }

    private UIElement CreateSectionHeader(string title, int completed, int total)
    {
      var banner = new Border
      {
        Background = new SolidColorBrush(Color.FromRgb(241, 248, 255)),
        BorderBrush = new SolidColorBrush(Color.FromRgb(3, 102, 214)),
        BorderThickness = new Thickness(3, 0, 0, 0),
        CornerRadius = new CornerRadius(0, 4, 4, 0),
        Padding = new Thickness(10, 6, 10, 6),
        Margin = new Thickness(0, 10, 0, 6)
      };

      var grid = new Grid();
      grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
      grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

      var titleBlock = new TextBlock
      {
        Text = title,
        FontSize = 13,
        FontWeight = FontWeights.Bold,
        Foreground = new SolidColorBrush(Color.FromRgb(36, 41, 46)),
        VerticalAlignment = VerticalAlignment.Center
      };
      Grid.SetColumn(titleBlock, 0);
      grid.Children.Add(titleBlock);

      if (total > 0)
      {
        var badge = new TextBlock
        {
          Text = $"{completed}/{total}",
          FontSize = 11,
          FontWeight = FontWeights.SemiBold,
          Foreground = completed == total
            ? new SolidColorBrush(Color.FromRgb(40, 167, 69))
            : new SolidColorBrush(Color.FromRgb(88, 96, 105)),
          VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(badge, 1);
        grid.Children.Add(badge);
      }

      banner.Child = grid;
      return banner;
    }

    private UIElement CreateItemRow(string checklistId, ChecklistItem item, bool isChecked)
    {
      var border = new Border
      {
        Background = isChecked
          ? new SolidColorBrush(Color.FromArgb(15, 40, 167, 69))
          : new SolidColorBrush(Color.FromRgb(255, 255, 255)),
        BorderBrush = new SolidColorBrush(Color.FromRgb(235, 238, 242)),
        BorderThickness = new Thickness(1),
        CornerRadius = new CornerRadius(4),
        Padding = new Thickness(8, 7, 8, 7),
        Margin = new Thickness(0, 2, 0, 2),
        Cursor = Cursors.Hand
      };

      var grid = new Grid();
      grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
      grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

      var checkBox = new CheckBox
      {
        IsChecked = isChecked,
        VerticalAlignment = VerticalAlignment.Top,
        Margin = new Thickness(0, 2, 8, 0),
        Tag = Tuple.Create(checklistId, item.Id)
      };

      var textBlock = new TextBlock
      {
        Text = item.Text,
        TextWrapping = TextWrapping.Wrap,
        FontSize = 12,
        LineHeight = 18,
        Foreground = isChecked
          ? new SolidColorBrush(Color.FromRgb(106, 115, 125))
          : new SolidColorBrush(Color.FromRgb(36, 41, 46))
      };

      if (isChecked)
      {
        textBlock.TextDecorations = TextDecorations.Strikethrough;
      }

      Grid.SetColumn(checkBox, 0);
      Grid.SetColumn(textBlock, 1);
      grid.Children.Add(checkBox);
      grid.Children.Add(textBlock);
      border.Child = grid;

      // Handle click on row to toggle checkbox
      border.MouseLeftButtonUp += (sender, e) =>
      {
        if (e.OriginalSource != checkBox)
        {
          checkBox.IsChecked = !checkBox.IsChecked;
        }
      };

      // Handle hover effect
      border.MouseEnter += (sender, e) =>
      {
        border.Background = isChecked
          ? new SolidColorBrush(Color.FromArgb(30, 40, 167, 69))
          : new SolidColorBrush(Color.FromRgb(246, 248, 250));
      };
      border.MouseLeave += (sender, e) =>
      {
        border.Background = isChecked
          ? new SolidColorBrush(Color.FromArgb(15, 40, 167, 69))
          : new SolidColorBrush(Color.FromRgb(255, 255, 255));
      };

      checkBox.Checked += (sender, e) => OnItemCheckedStateChanged(checklistId, item.Id, true);
      checkBox.Unchecked += (sender, e) => OnItemCheckedStateChanged(checklistId, item.Id, false);

      return border;
    }

    private void OnItemCheckedStateChanged(string checklistId, string itemId, bool isChecked)
    {
      if (_isUpdatingUi)
      {
        return;
      }

      var entry = GetOrCreateEntry(checklistId);
      entry.CompletedItems = entry.CompletedItems ?? new List<string>();

      if (isChecked)
      {
        if (!entry.CompletedItems.Any(id => string.Equals(id, itemId, StringComparison.OrdinalIgnoreCase)))
        {
          entry.CompletedItems.Add(itemId);
        }
      }
      else
      {
        entry.CompletedItems.RemoveAll(id => string.Equals(id, itemId, StringComparison.OrdinalIgnoreCase));
      }

      entry.LastModifiedUtc = DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture);

      // Auto-save immediately to folder's checklist_state.json
      SaveStateToDisk();

      // Refresh view
      RenderItems();
    }

    private void SaveStateToDisk()
    {
      if (string.IsNullOrWhiteSpace(_folderPath))
      {
        SetStatus("Cannot save: drawing folder path is unknown.", isError: true);
        return;
      }

      try
      {
        if (!string.IsNullOrWhiteSpace(_selectedChecklistId))
        {
          _state.LastActiveChecklistId = _selectedChecklistId;
        }

        ProjectChecklistStore.SaveFolderState(_folderPath, _state);
        SetStatus($"Saved to checklist_state.json ({DateTime.Now:hh:mm:ss tt})", isError: false);
      }
      catch (Exception ex)
      {
        SetStatus($"Save failed: {ex.Message}", isError: true);
      }
    }

    private void SetStatus(string message, bool isError)
    {
      StatusTextBlock.Text = message;
      if (isError)
      {
        StatusIcon.Text = "⚠";
        StatusIcon.Foreground = new SolidColorBrush(Color.FromRgb(220, 53, 69));
        StatusTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(220, 53, 69));
      }
      else
      {
        StatusIcon.Text = "✔";
        StatusIcon.Foreground = new SolidColorBrush(Color.FromRgb(40, 167, 69));
        StatusTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(88, 96, 105));
      }
    }

    private void ChecklistComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
      if (_isUpdatingUi)
      {
        return;
      }

      if (ChecklistComboBox.SelectedItem is ChecklistComboItem item)
      {
        _selectedChecklistId = item.Definition.Id;
        _state.LastActiveChecklistId = item.Definition.Id;
        SaveStateToDisk();
        RenderItems();
      }
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
      string query = SearchTextBox.Text ?? string.Empty;
      SearchPlaceholder.Visibility = string.IsNullOrEmpty(query) ? Visibility.Visible : Visibility.Collapsed;
      ClearSearchButton.Visibility = string.IsNullOrEmpty(query) ? Visibility.Collapsed : Visibility.Visible;
      RenderItems();
    }

    private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
    {
      SearchTextBox.Text = string.Empty;
      SearchTextBox.Focus();
    }

    private void HideCompletedCheckBox_Changed(object sender, RoutedEventArgs e)
    {
      RenderItems();
    }

    private void ResetChecklistButton_Click(object sender, RoutedEventArgs e)
    {
      var def = GetSelectedDefinition();
      if (def == null)
      {
        return;
      }

      var result = MessageBox.Show(
        $"Are you sure you want to uncheck all items for \"{def.Name}\"?\n\nThis will reset progress for this checklist in the current folder.",
        "Reset Checklist Progress",
        MessageBoxButton.YesNo,
        MessageBoxImage.Warning
      );

      if (result == MessageBoxResult.Yes)
      {
        var entry = GetOrCreateEntry(def.Id);
        entry.CompletedItems.Clear();
        entry.LastModifiedUtc = DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture);
        SaveStateToDisk();
        RenderItems();
      }
    }

    private void ReloadButton_Click(object sender, RoutedEventArgs e)
    {
      try
      {
        _state = ProjectChecklistStore.LoadFolderState(_folderPath);
        _definitions = ProjectChecklistStore.LoadAllChecklists(_folderPath);
        PopulateChecklists();
        SetStatus($"Reloaded from checklist_state.json ({DateTime.Now:hh:mm:ss tt})", isError: false);
      }
      catch (Exception ex)
      {
        SetStatus($"Reload failed: {ex.Message}", isError: true);
      }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
      SaveStateToDisk();
      Close();
    }

    protected override void OnClosed(EventArgs e)
    {
      SaveStateToDisk();
      base.OnClosed(e);
    }

    private sealed class ChecklistComboItem
    {
      internal ChecklistComboItem(ChecklistDefinition definition)
      {
        Definition = definition;
      }

      internal ChecklistDefinition Definition { get; }

      public override string ToString()
      {
        return Definition?.Name ?? "Untitled Checklist";
      }
    }
  }
}
