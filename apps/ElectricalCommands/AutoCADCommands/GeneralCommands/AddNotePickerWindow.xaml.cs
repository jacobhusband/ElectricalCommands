using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace ElectricalCommands
{
  public partial class AddNotePickerWindow : Window
  {
    private readonly List<AddNoteTab> _tabs;
    private List<AddNoteItem> _filteredNotes = new List<AddNoteItem>();

    public string SelectedNoteText { get; private set; }

    internal AddNotePickerWindow(List<AddNoteTab> tabs)
    {
      InitializeComponent();
      _tabs = tabs ?? new List<AddNoteTab>();
      TabComboBox.ItemsSource = _tabs;
      if (_tabs.Count > 0)
      {
        TabComboBox.SelectedIndex = 0;
      }

      SearchTextBox.Focus();
      RefreshNotes();
    }

    private void TabComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
      RefreshNotes();
    }

    private void SearchTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
      RefreshNotes();
    }

    private void NotesListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
      if (NotesListBox.SelectedItem != null)
      {
        ConfirmSelection();
      }
    }

    private void OK_Click(object sender, RoutedEventArgs e)
    {
      ConfirmSelection();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
      DialogResult = false;
      Close();
    }

    private void RefreshNotes()
    {
      AddNoteTab selectedTab = TabComboBox.SelectedItem as AddNoteTab;
      string query = SearchTextBox?.Text ?? string.Empty;
      string[] words = query
        .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(word => word.Trim())
        .Where(word => !string.IsNullOrWhiteSpace(word))
        .ToArray();

      IEnumerable<AddNoteItem> notes = selectedTab?.Notes ?? Enumerable.Empty<AddNoteItem>();
      if (words.Length > 0)
      {
        notes = notes.Where(note =>
        {
          string text = note?.Text ?? string.Empty;
          return words.All(word =>
            text.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0
          );
        });
      }

      _filteredNotes = notes.ToList();
      NotesListBox.ItemsSource = _filteredNotes;
      NotesListBox.SelectedIndex = _filteredNotes.Count > 0 ? 0 : -1;
      StatusTextBlock.Text = BuildStatusText(selectedTab, _filteredNotes.Count);
    }

    private string BuildStatusText(AddNoteTab selectedTab, int filteredCount)
    {
      if (selectedTab == null)
      {
        return "No tabs available.";
      }

      int totalCount = selectedTab.Notes?.Count ?? 0;
      return string.IsNullOrWhiteSpace(SearchTextBox?.Text)
        ? $"{totalCount} note(s)"
        : $"{filteredCount} of {totalCount} note(s)";
    }

    private void ConfirmSelection()
    {
      AddNoteItem selected = NotesListBox.SelectedItem as AddNoteItem;
      if (selected == null || string.IsNullOrWhiteSpace(selected.Text))
      {
        MessageBox.Show(
          "Please select a note.",
          "No Selection",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
        );
        return;
      }

      SelectedNoteText = selected.Text;
      DialogResult = true;
      Close();
    }
  }
}
