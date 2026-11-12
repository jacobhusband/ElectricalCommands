using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ElectricalCommands
{
  public enum NoteTableWindowAction
  {
    Cancel,
    Save,
    Create,
    UpdateTable,
    ImportText,
    ImportTable,
  }

  /// <summary>
  /// The NOTETABLE editor. It only edits the session; the command does the
  /// drawing work (imports, placing tables) between showings of the window.
  /// </summary>
  public partial class NoteTableWindow : Window
  {
    private readonly NoteTableSession _session;
    private NoteTableNoteItem _noteToFocus;

    internal NoteTableWindow(NoteTableSession session)
    {
      _session = session;
      InitializeComponent();

      SubtitleText.Text = string.IsNullOrEmpty(session.ProjectFilePath)
        ? "This drawing has never been saved, so the tabs can't be kept. " +
          "Save the drawing to share them with the project folder."
        : $"Tabs are shared by every drawing in this folder ({session.ProjectFilePath}).";
      SaveButton.IsEnabled = !string.IsNullOrEmpty(session.ProjectFilePath);
      NoteTabs.ItemsSource = session.Tabs;
      TextHeightBox.Text = FormatLength(session.TextHeight);
      TableWidthBox.Text = FormatLength(session.TableWidth);
      ScaleHintText.Text = session.ScaleHint ?? string.Empty;
      BoundaryPlacementButton.IsChecked = session.UseBoundary;
      PointPlacementButton.IsChecked = !session.UseBoundary;
      EraseImportedTextBox.IsChecked = session.EraseImportedText;
      ReplaceImportedTableBox.IsChecked = session.ReplaceImportedTable;
      StatusText.Text = session.StatusMessage ?? string.Empty;

      if (session.WindowBounds.HasValue)
      {
        Rect bounds = session.WindowBounds.Value;
        WindowStartupLocation = WindowStartupLocation.Manual;
        Left = bounds.Left;
        Top = bounds.Top;
        Width = bounds.Width;
        Height = bounds.Height;
      }
    }

    /// <summary>What the user asked for when the window closed.</summary>
    public NoteTableWindowAction Action { get; private set; } = NoteTableWindowAction.Cancel;

    public IReadOnlyList<NoteTableKindChoice> KindChoices { get; } = new[]
    {
      new NoteTableKindChoice(NoteTableKind.Keyed, "Keyed (hexagons)"),
      new NoteTableKindChoice(NoteTableKind.General, "General (1. 2. 3.)"),
    };

    private NoteTableTabItem SelectedTab => NoteTabs.SelectedItem as NoteTableTabItem;

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      int index = Math.Max(0, Math.Min(_session.SelectedTabIndex, _session.Tabs.Count - 1));
      if (_session.Tabs.Count > 0)
      {
        NoteTabs.SelectedIndex = index;
      }
      UpdateImportOptions();
    }

    private void Window_Closing(object sender, CancelEventArgs e)
    {
      _session.WindowBounds = new Rect(Left, Top, Width, Height);
      _session.SelectedTabIndex = Math.Max(0, NoteTabs.SelectedIndex);
    }

    private void NoteTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
      if (e.OriginalSource == NoteTabs)
      {
        UpdateImportOptions();
      }
    }

    private void UpdateImportOptions()
    {
      NoteTableTabItem tab = SelectedTab;
      EraseImportedTextBox.Visibility = tab != null && tab.HasImportedText
        ? Visibility.Visible
        : Visibility.Collapsed;
      ReplaceImportedTableBox.Visibility = tab != null && tab.HasImportedTable
        ? Visibility.Visible
        : Visibility.Collapsed;
    }

    private void AddTab_Click(object sender, RoutedEventArgs e)
    {
      string name = "New Tab";
      for (int suffix = 2; _session.Tabs.Any(tab => SameName(tab.Name, name)); suffix++)
      {
        name = $"New Tab {suffix}";
      }
      var tab = new NoteTableTabItem(name, NoteTableKind.Keyed, 1, Array.Empty<string>());
      _session.Tabs.Add(tab);
      NoteTabs.SelectedItem = tab;
    }

    private void DeleteTab_Click(object sender, RoutedEventArgs e)
    {
      NoteTableTabItem tab = SelectedTab;
      if (tab == null)
      {
        return;
      }
      if (_session.Tabs.Count == 1)
      {
        ShowError("Keep at least one tab. Rename this one instead.");
        return;
      }
      if (tab.Notes.Any(note => !string.IsNullOrWhiteSpace(note.Text)) &&
          MessageBox.Show(
            this,
            $"Delete the \"{tab.Name}\" tab and its {tab.Notes.Count} note(s)? " +
            "Other drawings in this folder share this tab.",
            "Delete Tab",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning) != MessageBoxResult.Yes)
      {
        return;
      }
      int index = NoteTabs.SelectedIndex;
      _session.Tabs.Remove(tab);
      NoteTabs.SelectedIndex = Math.Min(index, _session.Tabs.Count - 1);
    }

    private void AddNote_Click(object sender, RoutedEventArgs e)
    {
      NoteTableTabItem tab = SelectedTab;
      if (tab == null)
      {
        return;
      }
      var note = new NoteTableNoteItem(string.Empty);
      _noteToFocus = note;
      tab.Notes.Add(note);
    }

    private void NoteText_Loaded(object sender, RoutedEventArgs e)
    {
      if (sender is TextBox box && _noteToFocus != null &&
          ReferenceEquals(box.DataContext, _noteToFocus))
      {
        _noteToFocus = null;
        box.Focus();
        box.BringIntoView();
      }
    }

    private void MoveNoteUp_Click(object sender, RoutedEventArgs e)
    {
      MoveNote(sender, -1);
    }

    private void MoveNoteDown_Click(object sender, RoutedEventArgs e)
    {
      MoveNote(sender, 1);
    }

    private void MoveNote(object sender, int offset)
    {
      NoteTableTabItem tab = SelectedTab;
      if (tab == null || !((sender as FrameworkElement)?.DataContext is NoteTableNoteItem note))
      {
        return;
      }
      int index = tab.Notes.IndexOf(note);
      int target = index + offset;
      if (index >= 0 && target >= 0 && target < tab.Notes.Count)
      {
        tab.Notes.Move(index, target);
      }
    }

    private void RemoveNote_Click(object sender, RoutedEventArgs e)
    {
      if ((sender as FrameworkElement)?.DataContext is NoteTableNoteItem note)
      {
        SelectedTab?.Notes.Remove(note);
      }
    }

    private void ImportText_Click(object sender, RoutedEventArgs e)
    {
      CloseWith(NoteTableWindowAction.ImportText);
    }

    private void ImportTable_Click(object sender, RoutedEventArgs e)
    {
      CloseWith(NoteTableWindowAction.ImportTable);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
      CloseWith(NoteTableWindowAction.Save);
    }

    private void Create_Click(object sender, RoutedEventArgs e)
    {
      CloseWith(NoteTableWindowAction.Create);
    }

    private void UpdateTable_Click(object sender, RoutedEventArgs e)
    {
      CloseWith(NoteTableWindowAction.UpdateTable);
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
      Action = NoteTableWindowAction.Cancel;
      DialogResult = false;
      Close();
    }

    private void CloseWith(NoteTableWindowAction action)
    {
      if (!TryReadFields(action))
      {
        return;
      }
      Action = action;
      _session.StatusMessage = null;
      DialogResult = true;
      Close();
    }

    private bool TryReadFields(NoteTableWindowAction action)
    {
      if (SelectedTab == null)
      {
        ShowError("Add a tab first.");
        return false;
      }

      foreach (NoteTableTabItem tab in _session.Tabs)
      {
        if (string.IsNullOrWhiteSpace(tab.Name))
        {
          NoteTabs.SelectedItem = tab;
          ShowError("Every tab needs a name.");
          return false;
        }
        if (_session.Tabs.Any(other => !ReferenceEquals(other, tab) && SameName(other.Name, tab.Name)))
        {
          NoteTabs.SelectedItem = tab;
          ShowError($"Two tabs are named \"{tab.Name.Trim()}\". Tab names must be different.");
          return false;
        }
        if (!tab.HasValidFirstNumber)
        {
          NoteTabs.SelectedItem = tab;
          ShowError("\"Start at\" must be a whole number of 1 or more.");
          return false;
        }
      }

      bool creating = action == NoteTableWindowAction.Create;
      if (!TryParseLength(TextHeightBox.Text, out double textHeight))
      {
        if (creating)
        {
          ShowError("Enter a text height greater than zero.");
          TextHeightBox.Focus();
          return false;
        }
        textHeight = _session.TextHeight;
      }

      bool useBoundary = BoundaryPlacementButton.IsChecked == true;
      if (!TryParseLength(TableWidthBox.Text, out double tableWidth))
      {
        if (creating && !useBoundary)
        {
          ShowError("Enter a table width greater than zero.");
          TableWidthBox.Focus();
          return false;
        }
        tableWidth = _session.TableWidth;
      }

      if ((creating || action == NoteTableWindowAction.UpdateTable) &&
          !SelectedTab.Notes.Any(note => !string.IsNullOrWhiteSpace(note.Text)))
      {
        ShowError($"\"{SelectedTab.Name.Trim()}\" has no notes to put in a table.");
        return false;
      }

      _session.TextHeight = textHeight;
      _session.TableWidth = tableWidth;
      _session.UseBoundary = useBoundary;
      _session.EraseImportedText = EraseImportedTextBox.IsChecked == true;
      _session.ReplaceImportedTable = ReplaceImportedTableBox.IsChecked == true;
      _session.SelectedTabIndex = NoteTabs.SelectedIndex;
      return true;
    }

    private void ShowError(string message)
    {
      StatusText.Text = message;
    }

    private static bool SameName(string first, string second)
    {
      return string.Equals(
        (first ?? string.Empty).Trim(),
        (second ?? string.Empty).Trim(),
        StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseLength(string text, out double value)
    {
      return double.TryParse(
          (text ?? string.Empty).Trim(),
          NumberStyles.Float,
          CultureInfo.InvariantCulture,
          out value) &&
        value > 0.0 &&
        !double.IsInfinity(value);
    }

    private static string FormatLength(double value)
    {
      return value > 0.0
        ? value.ToString("0.#####", CultureInfo.InvariantCulture)
        : string.Empty;
    }
  }

  /// <summary>Everything NOTETABLE keeps between showings of its window.</summary>
  internal sealed class NoteTableSession
  {
    internal string ProjectFilePath { get; set; }
    internal DateTime LoadedFileTimeUtc { get; set; }
    internal ObservableCollection<NoteTableTabItem> Tabs { get; } =
      new ObservableCollection<NoteTableTabItem>();
    internal int SelectedTabIndex { get; set; }
    internal double TextHeight { get; set; }
    internal double TableWidth { get; set; }
    internal string ScaleHint { get; set; }
    internal bool UseBoundary { get; set; } = true;
    internal bool EraseImportedText { get; set; }
    internal bool ReplaceImportedTable { get; set; } = true;
    internal string StatusMessage { get; set; }
    internal Rect? WindowBounds { get; set; }
  }

  public sealed class NoteTableKindChoice
  {
    internal NoteTableKindChoice(NoteTableKind kind, string label)
    {
      Kind = kind;
      Label = label;
    }

    public NoteTableKind Kind { get; }
    public string Label { get; }
  }

  public sealed class NoteTableTabItem : INotifyPropertyChanged
  {
    private string _name;
    private NoteTableKind _kind;
    private string _firstNumberText;
    private string _importSummary = string.Empty;

    internal NoteTableTabItem(
      string name,
      NoteTableKind kind,
      int firstNumber,
      IEnumerable<string> notes)
    {
      _name = name ?? string.Empty;
      _kind = kind;
      _firstNumberText = Math.Max(1, firstNumber).ToString(CultureInfo.InvariantCulture);
      Notes = new ObservableCollection<NoteTableNoteItem>(
        (notes ?? Enumerable.Empty<string>()).Select(text => new NoteTableNoteItem(text)));
      Notes.CollectionChanged += Notes_CollectionChanged;
      Renumber();
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public ObservableCollection<NoteTableNoteItem> Notes { get; }

    public string Name
    {
      get => _name;
      set
      {
        _name = value ?? string.Empty;
        OnPropertyChanged(nameof(Name));
        OnPropertyChanged(nameof(Header));
      }
    }

    public string Header => string.IsNullOrWhiteSpace(_name) ? "(unnamed)" : _name.Trim();

    public NoteTableKind Kind
    {
      get => _kind;
      set
      {
        _kind = value;
        OnPropertyChanged(nameof(Kind));
        Renumber();
      }
    }

    public string FirstNumberText
    {
      get => _firstNumberText;
      set
      {
        _firstNumberText = value ?? string.Empty;
        OnPropertyChanged(nameof(FirstNumberText));
        Renumber();
      }
    }

    public bool IsEmpty => Notes.Count == 0;

    /// <summary>What was imported into this tab during this run, if anything.</summary>
    public string ImportSummary
    {
      get => _importSummary;
      internal set
      {
        _importSummary = value ?? string.Empty;
        OnPropertyChanged(nameof(ImportSummary));
      }
    }

    internal bool HasValidFirstNumber =>
      int.TryParse(_firstNumberText?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int number) &&
      number >= 1;

    internal int FirstNumber => HasValidFirstNumber
      ? int.Parse(_firstNumberText.Trim(), CultureInfo.InvariantCulture)
      : 1;

    /// <summary>Drawing objects this tab's notes came from during this run.</summary>
    internal NoteTableImport Import { get; set; }

    // Kept apart from Import so the window never touches AutoCAD types.
    internal bool HasImportedText { get; set; }
    internal bool HasImportedTable { get; set; }

    internal IReadOnlyList<string> GetNoteTexts()
    {
      return Notes
        .Select(note => NoteTableStore.NormalizeNoteText(note.Text))
        .Where(text => text.Length > 0)
        .ToList();
    }

    internal void AddNotes(IEnumerable<string> texts)
    {
      // An untouched blank note left by + Add Note would otherwise sit above
      // the imported ones.
      foreach (NoteTableNoteItem blank in Notes.Where(note => string.IsNullOrWhiteSpace(note.Text)).ToList())
      {
        Notes.Remove(blank);
      }
      foreach (string text in texts)
      {
        Notes.Add(new NoteTableNoteItem(text));
      }
    }

    private void Notes_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
    {
      OnPropertyChanged(nameof(IsEmpty));
      Renumber();
    }

    private void Renumber()
    {
      int number = FirstNumber;
      string suffix = _kind == NoteTableKind.General ? "." : string.Empty;
      foreach (NoteTableNoteItem note in Notes)
      {
        note.NumberText = number.ToString(CultureInfo.InvariantCulture) + suffix;
        number++;
      }
    }

    private void OnPropertyChanged(string propertyName)
    {
      PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
  }

  public sealed class NoteTableNoteItem : INotifyPropertyChanged
  {
    private string _text;
    private string _numberText = string.Empty;

    internal NoteTableNoteItem(string text)
    {
      _text = text ?? string.Empty;
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public string Text
    {
      get => _text;
      set
      {
        _text = value ?? string.Empty;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Text)));
      }
    }

    public string NumberText
    {
      get => _numberText;
      internal set
      {
        _numberText = value ?? string.Empty;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NumberText)));
      }
    }
  }
}
