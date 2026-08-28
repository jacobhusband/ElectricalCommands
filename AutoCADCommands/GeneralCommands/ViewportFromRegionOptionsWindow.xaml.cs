using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ElectricalCommands
{
  public partial class ViewportFromRegionOptionsWindow : Window
  {
    private readonly List<ViewportScaleOption> _scaleOptions;
    private readonly double _regionWidth;
    private readonly double _regionHeight;
    private readonly double _autoFitMaximumWidth;
    private readonly double _autoFitMaximumHeight;
    private ViewportScaleOption _resolvedScaleOption;

    internal ViewportFromRegionOptionsWindow(
      IEnumerable<string> layoutNames,
      IEnumerable<ViewportScaleOption> scaleOptions,
      string currentLayoutName,
      double regionWidth,
      double regionHeight,
      double autoFitMaximumWidth,
      double autoFitMaximumHeight
    )
    {
      InitializeComponent();

      List<string> layouts = (layoutNames ?? Enumerable.Empty<string>()).ToList();
      _scaleOptions = (scaleOptions ?? Enumerable.Empty<ViewportScaleOption>()).ToList();
      _regionWidth = regionWidth;
      _regionHeight = regionHeight;
      _autoFitMaximumWidth = autoFitMaximumWidth;
      _autoFitMaximumHeight = autoFitMaximumHeight;
      CustomScaleTextBox.Text = "1\" = 10'-0\"";

      LayoutComboBox.ItemsSource = layouts;
      int currentLayoutIndex = layouts.FindIndex(layout =>
        layout.Equals(currentLayoutName, StringComparison.OrdinalIgnoreCase)
      );
      LayoutComboBox.SelectedIndex = currentLayoutIndex >= 0 ? currentLayoutIndex : 0;

      ScaleComboBox.ItemsSource = _scaleOptions;
      ScaleComboBox.SelectedIndex = _scaleOptions.Count > 0 ? 0 : -1;

      RegionSizeTextBlock.Text =
        $"Selected model region: {_regionWidth:0.##} x {_regionHeight:0.##} drawing units";
      RefreshScaleSummary();
    }

    internal string SelectedLayoutName { get; private set; }
    internal ViewportScaleOption SelectedScaleOption { get; private set; }

    private void ScaleComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
      RefreshScaleSummary();
    }

    private void CustomScaleTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
      RefreshScaleSummary();
    }

    private void RefreshScaleSummary()
    {
      if (ScaleSummaryTextBlock == null ||
          CreateButton == null ||
          CustomScaleGrid == null ||
          _scaleOptions == null)
      {
        return;
      }

      ViewportScaleOption requestedScale = ScaleComboBox.SelectedItem as ViewportScaleOption;
      bool isCustomScale = requestedScale != null && requestedScale.IsCustom;
      CustomScaleGrid.Visibility = isCustomScale ? Visibility.Visible : Visibility.Collapsed;

      string validationMessage = string.Empty;
      ViewportScaleOption resolvedScale;
      if (isCustomScale)
      {
        ViewportScaleOption.TryCreateCustom(
          CustomScaleTextBox.Text,
          out resolvedScale,
          out validationMessage
        );
      }
      else
      {
        resolvedScale = GeneralCommands.ResolveScaleOption(
          requestedScale,
          _scaleOptions,
          _regionWidth,
          _regionHeight
        );
      }

      _resolvedScaleOption = resolvedScale;

      if (resolvedScale == null)
      {
        ScaleSummaryTextBlock.Text = isCustomScale
          ? validationMessage
          : $"Auto Fit cannot fit this region within {_autoFitMaximumWidth:0.##} x " +
            $"{_autoFitMaximumHeight:0.##}. Choose a specific scale.";
        ScaleSummaryTextBlock.Foreground = Brushes.Firebrick;
        CreateButton.IsEnabled = false;
        return;
      }

      double viewportWidth = _regionWidth / resolvedScale.ModelUnitsPerPaperUnit;
      double viewportHeight = _regionHeight / resolvedScale.ModelUnitsPerPaperUnit;
      string prefix = requestedScale != null && requestedScale.IsAutoFit
        ? $"Auto Fit selects {resolvedScale.DisplayName}. "
        : string.Empty;

      ScaleSummaryTextBlock.Text =
        $"{prefix}Viewport size: {viewportWidth:0.##} x {viewportHeight:0.##} paperspace units.";
      ScaleSummaryTextBlock.Foreground = Brushes.Black;
      CreateButton.IsEnabled = true;
    }

    private void Create_Click(object sender, RoutedEventArgs e)
    {
      string selectedLayout = LayoutComboBox.SelectedItem as string;
      if (string.IsNullOrWhiteSpace(selectedLayout) || _resolvedScaleOption == null)
      {
        MessageBox.Show(
          "Select a sheet/layout and enter a valid viewport scale.",
          "Viewport Options",
          MessageBoxButton.OK,
          MessageBoxImage.Warning
        );
        return;
      }

      SelectedLayoutName = selectedLayout;
      SelectedScaleOption = _resolvedScaleOption;
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
