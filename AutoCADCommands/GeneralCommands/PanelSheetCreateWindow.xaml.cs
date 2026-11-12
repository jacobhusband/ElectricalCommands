using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ElectricalCommands
{
  public partial class PanelSheetCreateWindow : Window
  {
    private static readonly string[] VoltageChoices =
      { "120/208V", "120/240V", "277/480V", "208Y/120V", "480Y/277V" };
    private static readonly string[] BusRatingChoices =
      { "100A", "125A", "150A", "200A", "225A", "400A", "600A", "800A" };
    private static readonly string[] MainChoices =
      { "MLO", "100A MCB", "125A MCB", "200A MCB", "225A MCB", "400A MCB" };
    private static readonly string[] WireChoices = { "3", "4" };
    private static readonly string[] PhaseChoices = { "1", "3" };
    private static readonly string[] TypeChoices =
      { "NEMA 1", "NEMA 3R", "NEMA 4X" };
    private static readonly string[] MountingChoices = { "SURFACE", "RECESSED" };
    private static readonly string[] AicChoices =
      { "EXISTING", "10KAIC", "14KAIC", "22KAIC", "42KAIC", "65KAIC" };
    private static readonly string[] ConditionChoices =
      { "NEW", "EXISTING", "RELOCATED" };
    private static readonly string[] CircuitCountChoices =
      { "6", "8", "12", "16", "18", "20", "24", "30", "36", "42" };

    private readonly string _panelName;
    private readonly PanelTemplateCatalog _catalog;
    private bool _loadingTemplate;
    private bool _breakerPhasesEdited;
    private bool _updatingBreakerPhases;

    /// <summary>The worksheet to create; null unless the user clicked Create.</summary>
    internal PanelSheetRequest Request { get; private set; }

    internal PanelSheetCreateWindow(
      string panelName,
      string workbookName,
      PanelTemplateCatalog catalog)
    {
      _panelName = panelName ?? string.Empty;
      _catalog = catalog ?? new PanelTemplateCatalog();
      InitializeComponent();

      SubtitleText.Text =
        $"Panel \"{_panelName}\" is not in {workbookName}. Create a worksheet " +
        "for it by copying one of the workbook's panel sheets.";
      SheetNameBox.Text = PanelSheetRules.SuggestSheetName(_panelName);
      FillCombo(CircuitCountBox, CircuitCountChoices, string.Empty);

      PanelTemplateInfo preferred = _catalog.Templates.FirstOrDefault(
          t => string.Equals(
            t.SheetName,
            "TEMPLATE",
            StringComparison.OrdinalIgnoreCase)) ??
        _catalog.Templates.FirstOrDefault(
          t => t.SheetName.IndexOf(
            "template",
            StringComparison.OrdinalIgnoreCase) >= 0) ??
        _catalog.Templates.FirstOrDefault();

      foreach (PanelTemplateInfo template in _catalog.Templates)
      {
        TemplateBox.Items.Add(new ComboBoxItem
        {
          Content = template.DisplayText,
          Tag = template,
        });
      }
      foreach (ComboBoxItem item in TemplateBox.Items)
      {
        if (ReferenceEquals(item.Tag, preferred))
        {
          TemplateBox.SelectedItem = item;
          break;
        }
      }
    }

    private PanelTemplateInfo SelectedTemplate =>
      (TemplateBox.SelectedItem as ComboBoxItem)?.Tag as PanelTemplateInfo;

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      SheetNameBox.Focus();
      SheetNameBox.SelectAll();
    }

    private void TemplateBox_SelectionChanged(
      object sender,
      SelectionChangedEventArgs e)
    {
      PanelTemplateInfo template = SelectedTemplate;
      if (template == null)
      {
        return;
      }

      _loadingTemplate = true;
      try
      {
        PanelHeaderValues header = template.Header;
        FillCombo(VoltageBox, VoltageChoices, header.Voltage);
        FillCombo(BusRatingBox, BusRatingChoices, header.BusRating);
        FillCombo(MainBox, MainChoices, header.Main);
        FillCombo(WiresBox, WireChoices, header.Wires);
        FillCombo(PhaseBox, PhaseChoices, header.Phase);
        FillCombo(TypeBox, TypeChoices, header.Type);
        FillCombo(MountingBox, MountingChoices, header.Mounting);
        FillCombo(AicBox, AicChoices, header.AicRating);
        FillCombo(ConditionBox, ConditionChoices, header.Condition);
        CircuitCountBox.Text = template.CircuitCapacity.ToString();
        CircuitHintText.Text =
          $"\"{template.SheetName}\" has {template.CircuitCapacity} circuit positions, " +
          "the most this new sheet can have.";
        _breakerPhasesEdited = false;
        SetBreakerPhases(PanelSheetRules.DefaultBreakerPhases(header.Phase));
      }
      finally
      {
        _loadingTemplate = false;
      }
      StatusText.Text = string.Empty;
    }

    private void PhaseBox_SelectionChanged(
      object sender,
      SelectionChangedEventArgs e)
    {
      if (_loadingTemplate || _breakerPhasesEdited)
      {
        return;
      }
      SetBreakerPhases(
        PanelSheetRules.DefaultBreakerPhases(
          PhaseBox.SelectedItem as string));
    }

    private void BreakerPhasesBox_TextChanged(
      object sender,
      TextChangedEventArgs e)
    {
      if (!_updatingBreakerPhases)
      {
        _breakerPhasesEdited = true;
      }
    }

    private void SetBreakerPhases(string text)
    {
      _updatingBreakerPhases = true;
      try
      {
        BreakerPhasesBox.Text = text;
      }
      finally
      {
        _updatingBreakerPhases = false;
      }
    }

    private void Create_Click(object sender, RoutedEventArgs e)
    {
      PanelTemplateInfo template = SelectedTemplate;
      if (template == null)
      {
        ShowError("Choose a template sheet.");
        return;
      }

      string sheetName = SheetNameBox.Text.Trim();
      if (!PanelSheetRules.TryValidateSheetName(
        sheetName,
        _catalog.AllSheetNames,
        out string error))
      {
        ShowError(error);
        SheetNameBox.Focus();
        return;
      }

      if (!PanelSheetRules.TryParseCircuitCount(
        CircuitCountBox.Text,
        template.CircuitCapacity,
        out int circuitCount,
        out error))
      {
        ShowError(error);
        CircuitCountBox.Focus();
        return;
      }

      if (!PanelSheetRules.TryParseBreakerPhases(
        BreakerPhasesBox.Text,
        out string[] phases,
        out error))
      {
        ShowError(error);
        BreakerPhasesBox.Focus();
        return;
      }

      Request = new PanelSheetRequest
      {
        TemplateSheetName = template.SheetName,
        NewSheetName = sheetName,
        PanelName = _panelName,
        CircuitCount = circuitCount,
        BreakerPhases = phases,
        Header = new PanelHeaderValues
        {
          Voltage = VoltageBox.Text.Trim(),
          BusRating = BusRatingBox.Text.Trim(),
          Main = MainBox.Text.Trim(),
          Wires = WiresBox.SelectedItem as string ?? string.Empty,
          Phase = PhaseBox.SelectedItem as string ?? string.Empty,
          Type = TypeBox.Text.Trim(),
          Mounting = MountingBox.Text.Trim(),
          AicRating = AicBox.Text.Trim(),
          Condition = ConditionBox.SelectedItem as string ?? string.Empty,
        },
      };
      DialogResult = true;
      Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
      DialogResult = false;
      Close();
    }

    private void ShowError(string message)
    {
      StatusText.Text = message;
    }

    /// <summary>
    /// Loads the suggestions and selects (or, for editable boxes, types) the
    /// current value, adding it to the list if it is not already a choice.
    /// </summary>
    private static void FillCombo(
      ComboBox box,
      IEnumerable<string> choices,
      string current)
    {
      current = (current ?? string.Empty).Trim();
      List<string> items = choices.ToList();
      if (current.Length > 0 &&
          !items.Any(item => string.Equals(
            item,
            current,
            StringComparison.OrdinalIgnoreCase)))
      {
        items.Insert(0, current);
      }

      box.Items.Clear();
      foreach (string item in items)
      {
        box.Items.Add(item);
      }

      if (box.IsEditable)
      {
        box.Text = current;
      }
      else
      {
        box.SelectedItem = items.FirstOrDefault(item => string.Equals(
          item,
          current,
          StringComparison.OrdinalIgnoreCase));
      }
    }
  }
}
