using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ElectricalCommands
{
  internal sealed class DedicatedEquipmentPresetItem
  {
    internal DedicatedEquipmentPresetItem(DedicatedEquipmentLoad equipment)
    {
      Equipment = equipment ?? new DedicatedEquipmentLoad();
    }

    internal DedicatedEquipmentLoad Equipment { get; }

    public string Description => Equipment.Description ?? string.Empty;

    public string Summary =>
      $"{Equipment.Kva:0.###} kVA  |  {Equipment.Voltage}V  |  " +
      $"{Equipment.Poles}P  |  {ResolveBreakerText()}";

    public string AliasText => Equipment.Aliases != null &&
      Equipment.Aliases.Count > 0
        ? "Aliases: " + string.Join(", ", Equipment.Aliases)
        : "No aliases";

    internal string SearchText =>
      Description + " " + AliasText + " " + Summary;

    private string ResolveBreakerText()
    {
      return Equipment.MocpAmps.HasValue
        ? Equipment.MocpAmps.Value.ToString(
            CultureInfo.InvariantCulture) + "A"
        : "calculated breaker";
    }
  }

  public partial class DedicatedEquipmentPickerWindow : Window
  {
    private List<DedicatedEquipmentLoad> _presets =
      new List<DedicatedEquipmentLoad>();
    private List<DedicatedEquipmentPresetItem> _filteredPresets =
      new List<DedicatedEquipmentPresetItem>();
    private string _catalogPath = string.Empty;
    private string _editingOriginalDescription = string.Empty;
    private bool _oneTimeMode;
    private bool _suppressSelectionChange;

    internal DedicatedEquipmentLoad SelectedEquipment { get; private set; }

    internal DedicatedEquipmentPickerWindow()
    {
      InitializeComponent();
      PolesComboBox.ItemsSource = new[] { 1, 2, 3 };
      PolesComboBox.SelectedItem = 1;
      ReloadPresets();
      SearchTextBox.Focus();
    }

    private void ReloadPresets(string descriptionToSelect = null)
    {
      _presets = DedicatedEquipmentCatalog.LoadAll(out _catalogPath);
      CatalogPathTextBlock.Text = "Preset file: " + _catalogPath;
      CatalogPathTextBlock.ToolTip = _catalogPath;
      RefreshPresetList(descriptionToSelect);
    }

    private void SearchTextBox_TextChanged(
      object sender,
      TextChangedEventArgs e)
    {
      if (PresetListBox != null)
      {
        RefreshPresetList();
      }
    }

    private void RefreshPresetList(string descriptionToSelect = null)
    {
      if (PresetListBox == null || _presets == null)
      {
        return;
      }

      string[] words = (SearchTextBox?.Text ?? string.Empty)
        .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(word => word.Trim())
        .Where(word => word.Length > 0)
        .ToArray();

      IEnumerable<DedicatedEquipmentPresetItem> items = _presets
        .Select(preset => new DedicatedEquipmentPresetItem(preset));
      if (words.Length > 0)
      {
        items = items.Where(item => words.All(word =>
          item.SearchText.IndexOf(
            word,
            StringComparison.OrdinalIgnoreCase) >= 0));
      }

      _filteredPresets = items
        .OrderBy(item => item.Description, StringComparer.OrdinalIgnoreCase)
        .ToList();

      _suppressSelectionChange = true;
      try
      {
        PresetListBox.ItemsSource = _filteredPresets;
        int selectedIndex = -1;
        if (!string.IsNullOrWhiteSpace(descriptionToSelect))
        {
          selectedIndex = _filteredPresets.FindIndex(item =>
            string.Equals(
              item.Description,
              descriptionToSelect,
              StringComparison.OrdinalIgnoreCase));
        }
        if (selectedIndex < 0 && _filteredPresets.Count > 0)
        {
          selectedIndex = 0;
        }
        PresetListBox.SelectedIndex = selectedIndex;
      }
      finally
      {
        _suppressSelectionChange = false;
      }

      PresetCountTextBlock.Text = words.Length == 0
        ? $"{_presets.Count} preset(s)"
        : $"{_filteredPresets.Count} of {_presets.Count} preset(s)";

      if (PresetListBox.SelectedItem is DedicatedEquipmentPresetItem selected)
      {
        PopulateEditor(selected.Equipment);
      }
      else if (_filteredPresets.Count == 0)
      {
        BeginNewCircuit(true);
        SetStatus(
          words.Length > 0
            ? "No preset matches this search. Enter the remaining values to " +
              "use it once, or choose New Preset to save it."
            : "There are no saved presets. Use a one-time circuit or add a " +
              "new preset.",
          false);
      }
    }

    private void PresetListBox_SelectionChanged(
      object sender,
      SelectionChangedEventArgs e)
    {
      if (_suppressSelectionChange)
      {
        return;
      }

      if (PresetListBox.SelectedItem is DedicatedEquipmentPresetItem selected)
      {
        PopulateEditor(selected.Equipment);
      }
    }

    private void PresetListBox_MouseDoubleClick(
      object sender,
      MouseButtonEventArgs e)
    {
      if (PresetListBox.SelectedItem != null)
      {
        UseCircuit();
      }
    }

    private void PopulateEditor(DedicatedEquipmentLoad equipment)
    {
      if (equipment == null)
      {
        return;
      }

      _oneTimeMode = false;
      _editingOriginalDescription = equipment.Description ?? string.Empty;
      DescriptionTextBox.Text = equipment.Description ?? string.Empty;
      AliasesTextBox.Text = string.Join(
        ", ",
        equipment.Aliases ?? new List<string>());
      KvaTextBox.Text = equipment.Kva.ToString(
        "0.###",
        CultureInfo.CurrentCulture);
      VoltageTextBox.Text = equipment.Voltage.ToString(
        CultureInfo.CurrentCulture);
      PolesComboBox.SelectedItem = equipment.Poles;
      McaTextBox.Text = equipment.McaAmps.HasValue
        ? equipment.McaAmps.Value.ToString(
            "0.##",
            CultureInfo.CurrentCulture)
        : string.Empty;
      MocpTextBox.Text = equipment.MocpAmps.HasValue
        ? equipment.MocpAmps.Value.ToString(CultureInfo.CurrentCulture)
        : string.Empty;

      EditorTitleTextBlock.Text = "Preset details";
      EditorHelpTextBlock.Text =
        "Review the saved values, then use the circuit. Edit and save to " +
        "update this preset.";
      SavePresetButton.IsEnabled = true;
      SavePresetButton.Content = "Update Preset";
      RemovePresetButton.IsEnabled = true;
      ValidationTextBlock.Text =
        "This circuit will use the selected preset values.";
      ValidationTextBlock.Foreground = Brushes.SeaGreen;
    }

    private void NewPreset_Click(object sender, RoutedEventArgs e)
    {
      BeginNewCircuit(false);
    }

    private void NewOneTime_Click(object sender, RoutedEventArgs e)
    {
      BeginNewCircuit(true);
    }

    private void BeginNewCircuit(bool oneTime)
    {
      _suppressSelectionChange = true;
      try
      {
        PresetListBox.SelectedItem = null;
      }
      finally
      {
        _suppressSelectionChange = false;
      }

      _oneTimeMode = oneTime;
      _editingOriginalDescription = string.Empty;
      DescriptionTextBox.Text = NormalizeDescription(
        SearchTextBox?.Text ?? string.Empty);
      AliasesTextBox.Text = string.Empty;
      KvaTextBox.Text = string.Empty;
      VoltageTextBox.Text = "120";
      PolesComboBox.SelectedItem = 1;
      McaTextBox.Text = string.Empty;
      MocpTextBox.Text = string.Empty;
      RemovePresetButton.IsEnabled = false;
      SavePresetButton.IsEnabled = true;
      SavePresetButton.Content = oneTime
        ? "Save as Preset"
        : "Save Preset";
      EditorTitleTextBlock.Text = oneTime
        ? "One-time circuit"
        : "New preset";
      EditorHelpTextBlock.Text = oneTime
        ? "Enter circuit values. They will be used once and will not be " +
          "added to the preset list unless you choose Save as Preset."
        : "Enter circuit values, save the preset, then use it now or later.";
      ValidationTextBlock.Text = oneTime
        ? "One-time mode: this circuit will not be saved."
        : "Enter the new preset details.";
      ValidationTextBlock.Foreground = Brushes.SteelBlue;

      if (DescriptionTextBox.Text.Length == 0)
      {
        DescriptionTextBox.Focus();
      }
      else
      {
        KvaTextBox.Focus();
      }
    }

    private void SavePreset_Click(object sender, RoutedEventArgs e)
    {
      if (!TryBuildEquipment(out DedicatedEquipmentLoad equipment))
      {
        return;
      }

      try
      {
        DedicatedEquipmentCatalog.SaveOrUpdate(
          equipment,
          _editingOriginalDescription);
        SearchTextBox.Text = string.Empty;
        ReloadPresets(equipment.Description);
        SetStatus(
          $"Saved {equipment.Description}. It is selected and ready to use.",
          false);
      }
      catch (System.Exception ex)
      {
        SetStatus("Could not save the preset: " + ex.Message, true);
      }
    }

    private void RemovePreset_Click(object sender, RoutedEventArgs e)
    {
      if (!(PresetListBox.SelectedItem is DedicatedEquipmentPresetItem selected))
      {
        SetStatus("Select a preset to remove.", true);
        return;
      }

      MessageBoxResult confirmation = MessageBox.Show(
        this,
        $"Remove the {selected.Description} preset?\n\n" +
        "This does not affect circuits already added to drawings.",
        "Remove Dedicated-Circuit Preset",
        MessageBoxButton.YesNo,
        MessageBoxImage.Warning,
        MessageBoxResult.No);
      if (confirmation != MessageBoxResult.Yes)
      {
        return;
      }

      try
      {
        DedicatedEquipmentCatalog.Remove(selected.Description);
        ReloadPresets();
        SetStatus($"Removed {selected.Description}.", false);
      }
      catch (System.Exception ex)
      {
        SetStatus("Could not remove the preset: " + ex.Message, true);
      }
    }

    private void SuggestBreaker_Click(object sender, RoutedEventArgs e)
    {
      if (!TryReadElectricalValues(
        out double kva,
        out int voltage,
        out int poles,
        out double? mca))
      {
        return;
      }

      try
      {
        double calculatedAmps = GeneralCommands.CalculateDedicatedLoadAmps(
          kva,
          voltage,
          poles);
        int breaker = GeneralCommands.SelectStandardDedicatedBreaker(
          Math.Max(calculatedAmps, mca ?? 0.0));
        MocpTextBox.Text = breaker.ToString(CultureInfo.CurrentCulture);
        SetStatus(
          $"Suggested {breaker}A from {calculatedAmps:0.##}A connected " +
          "load" +
          (mca.HasValue ? $" and {mca.Value:0.##}A MCA." : "."),
          false);
      }
      catch (System.Exception ex)
      {
        SetStatus(ex.Message, true);
      }
    }

    private void UseCircuit_Click(object sender, RoutedEventArgs e)
    {
      UseCircuit();
    }

    private void UseCircuit()
    {
      if (!TryBuildEquipment(out DedicatedEquipmentLoad equipment))
      {
        return;
      }

      SelectedEquipment = equipment;
      DialogResult = true;
      Close();
    }

    private bool TryBuildEquipment(out DedicatedEquipmentLoad equipment)
    {
      equipment = null;
      string description = NormalizeDescription(DescriptionTextBox.Text);
      if (description.Length == 0)
      {
        SetStatus("Enter an equipment description.", true);
        DescriptionTextBox.Focus();
        return false;
      }

      if (!TryReadElectricalValues(
        out double kva,
        out int voltage,
        out int poles,
        out double? mca))
      {
        return false;
      }

      int mocp;
      if (string.IsNullOrWhiteSpace(MocpTextBox.Text))
      {
        try
        {
          mocp = GeneralCommands.SelectStandardDedicatedBreaker(
            Math.Max(
              GeneralCommands.CalculateDedicatedLoadAmps(
                kva,
                voltage,
                poles),
              mca ?? 0.0));
          MocpTextBox.Text = mocp.ToString(CultureInfo.CurrentCulture);
        }
        catch (System.Exception ex)
        {
          SetStatus(ex.Message, true);
          return false;
        }
      }
      else if (!TryParseInteger(MocpTextBox.Text, out mocp) ||
               mocp < 1 || mocp > 1200)
      {
        SetStatus("Breaker / MOCP must be between 1A and 1200A.", true);
        MocpTextBox.Focus();
        return false;
      }

      double calculatedAmps = GeneralCommands.CalculateDedicatedLoadAmps(
        kva,
        voltage,
        poles);
      double minimumBreakerAmps = Math.Max(calculatedAmps, mca ?? 0.0);
      if (mocp + 1e-9 < minimumBreakerAmps)
      {
        SetStatus(
          $"The {mocp}A breaker is below the required " +
          $"{minimumBreakerAmps:0.##}A load/MCA.",
          true);
        MocpTextBox.Focus();
        return false;
      }

      List<string> aliases = (AliasesTextBox.Text ?? string.Empty)
        .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(NormalizeDescription)
        .Where(alias => alias.Length > 0 &&
          !string.Equals(
            alias,
            description,
            StringComparison.OrdinalIgnoreCase))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

      equipment = new DedicatedEquipmentLoad
      {
        Description = description,
        Aliases = aliases,
        Kva = kva,
        Voltage = voltage,
        Poles = poles,
        McaAmps = mca,
        MocpAmps = mocp,
        LoadTypeCode = "D",
      };
      return true;
    }

    private bool TryReadElectricalValues(
      out double kva,
      out int voltage,
      out int poles,
      out double? mca)
    {
      kva = 0.0;
      voltage = 0;
      poles = 0;
      mca = null;

      if (!TryParseDouble(KvaTextBox.Text, out kva) ||
          kva <= 0.0 ||
          double.IsNaN(kva) ||
          double.IsInfinity(kva))
      {
        SetStatus("Connected load must be a positive kVA value.", true);
        KvaTextBox.Focus();
        return false;
      }
      if (!TryParseInteger(VoltageTextBox.Text, out voltage) ||
          voltage < 1 || voltage > 1000)
      {
        SetStatus("Voltage must be between 1V and 1000V.", true);
        VoltageTextBox.Focus();
        return false;
      }
      if (!(PolesComboBox.SelectedItem is int selectedPoles) ||
          selectedPoles < 1 || selectedPoles > 3)
      {
        SetStatus("Select one, two, or three poles.", true);
        PolesComboBox.Focus();
        return false;
      }
      poles = selectedPoles;

      if (!string.IsNullOrWhiteSpace(McaTextBox.Text))
      {
        if (!TryParseDouble(McaTextBox.Text, out double parsedMca) ||
            parsedMca <= 0.0 ||
            double.IsNaN(parsedMca) ||
            double.IsInfinity(parsedMca))
        {
          SetStatus("MCA must be a positive amp value or left blank.", true);
          McaTextBox.Focus();
          return false;
        }
        mca = parsedMca;
      }
      return true;
    }

    private static bool TryParseDouble(string text, out double value)
    {
      return double.TryParse(
          text,
          NumberStyles.Float,
          CultureInfo.CurrentCulture,
          out value) ||
        double.TryParse(
          text,
          NumberStyles.Float,
          CultureInfo.InvariantCulture,
          out value);
    }

    private static bool TryParseInteger(string text, out int value)
    {
      return int.TryParse(
          text,
          NumberStyles.Integer,
          CultureInfo.CurrentCulture,
          out value) ||
        int.TryParse(
          text,
          NumberStyles.Integer,
          CultureInfo.InvariantCulture,
          out value);
    }

    private static string NormalizeDescription(string text)
    {
      return Regex.Replace(
        text ?? string.Empty,
        @"\s+",
        " ").Trim().ToUpperInvariant();
    }

    private void SetStatus(string message, bool isError)
    {
      ValidationTextBlock.Text = message ?? string.Empty;
      ValidationTextBlock.Foreground = isError
        ? Brushes.Firebrick
        : Brushes.SeaGreen;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
      DialogResult = false;
      Close();
    }
  }
}
