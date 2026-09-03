using System;
using System.Text;
using System.Windows;
using System.Windows.Media;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ElectricalCommands
{
  public partial class SwitchConfigurationWindow : Window
  {
    private ProjectSwitchSettings _settings;
    private SwitchType _currentType = SwitchType.Standard;

    public SwitchConfigurationWindow()
    {
      InitializeComponent();
      Loaded += OnWindowLoaded;
    }

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc != null)
      {
        _settings = SwitchConfigurationStore.LoadSettings(doc.Database, doc);
        if (SwitchConfigurationStore.TryResolveSettingsPath(doc, doc.Database, out string settingsPath, out string projectRoot))
        {
          ProjectLocationText.Text = $"Project: {projectRoot}\nConfig: {settingsPath}";
        }
        else
        {
          ProjectLocationText.Text = "Project: (Drawing not in recognized project structure; using drawing-local/in-session settings)";
        }
      }
      else
      {
        _settings = new ProjectSwitchSettings();
        ProjectLocationText.Text = "Project: (No active drawing)";
      }

      RefreshUI();
    }

    private void OnSwitchTypeChanged(object sender, RoutedEventArgs e)
    {
      if (RadioStandard != null && RadioStandard.IsChecked == true)
      {
        _currentType = SwitchType.Standard;
      }
      else if (RadioDimmer != null && RadioDimmer.IsChecked == true)
      {
        _currentType = SwitchType.Dimmer;
      }
      else if (RadioOccupancy != null && RadioOccupancy.IsChecked == true)
      {
        _currentType = SwitchType.Occupancy;
      }

      RefreshUI();
    }

    private void RefreshUI()
    {
      if (_settings == null) return;
      var typeConfig = _settings.GetTypeConfig(_currentType);

      UpdateCard(NorthBadge, NorthStatusText, NorthDetailsText, typeConfig.North);
      UpdateCard(EastBadge, EastStatusText, EastDetailsText, typeConfig.East);
      UpdateCard(SouthBadge, SouthStatusText, SouthDetailsText, typeConfig.South);
      UpdateCard(WestBadge, WestStatusText, WestDetailsText, typeConfig.West);
    }

    private void UpdateCard(
      System.Windows.Controls.Border badge,
      System.Windows.Controls.TextBlock statusText,
      System.Windows.Controls.TextBlock detailsText,
      SwitchOrientationConfig cfg)
    {
      if (cfg != null && cfg.IsConfigured)
      {
        badge.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 245, 225));
        statusText.Text = "Configured";
        statusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(46, 125, 50));

        var sb = new StringBuilder();
        double rotDeg = cfg.Block.Rotation * 180.0 / Math.PI;
        sb.AppendLine($"Block: {cfg.Block.BlockName}");
        sb.AppendLine($"Rotation: {rotDeg:0}° | Layer: {cfg.Block.Layer}");
        sb.AppendLine($"Scale: ({cfg.Block.ScaleX:0.##}, {cfg.Block.ScaleY:0.##})");
        if (!string.IsNullOrWhiteSpace(cfg.Block.VisibilityState))
        {
          sb.AppendLine($"Visibility: {cfg.Block.VisibilityState}");
        }

        if (cfg.TextObjects != null && cfg.TextObjects.Count > 0)
        {
          sb.AppendLine($"\nText Objects ({cfg.TextObjects.Count}):");
          for (int i = 0; i < cfg.TextObjects.Count; i++)
          {
            var txt = cfg.TextObjects[i];
            sb.AppendLine($" • \"{txt.TextString}\" [Style: {txt.TextStyleName}, H: {txt.Height:0.###}]");
            sb.AppendLine($"   Offset: ({txt.RelativeOffset.X:0.##}, {txt.RelativeOffset.Y:0.##})");
          }
        }
        else
        {
          sb.AppendLine("\nNo associated text objects.");
        }

        detailsText.Text = sb.ToString();
      }
      else
      {
        badge.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(237, 242, 247));
        statusText.Text = "Not Configured";
        statusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(113, 128, 150));
        detailsText.Text = "No block or text configuration set for this orientation.";
      }
    }

    private void CaptureOrientation(SwitchOrientation orientation)
    {
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        MessageBox.Show("No active AutoCAD document found.", "ACIES", MessageBoxButton.OK, MessageBoxImage.Warning);
        return;
      }

      Editor ed = doc.Editor;
      Database db = doc.Database;

      EditorUserInteraction interaction = null;
      try
      {
        interaction = ed.StartUserInteraction(this);
        if (SwitchPlacementEngine.CaptureFromSelection(ed, db, out var capturedConfig, out string error))
        {
          var typeConfig = _settings.GetTypeConfig(_currentType);
          typeConfig.SetOrientation(orientation, capturedConfig);

          SwitchConfigurationStore.SaveSettings(db, _settings, out string saveErr, doc);
          ed.WriteMessage($"\nConfigured {orientation} for {typeConfig.DisplayName}.");
        }
        else
        {
          ed.WriteMessage($"\nCapture cancelled: {error}");
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nCapture error: {ex.Message}");
      }
      finally
      {
        interaction?.End();
        Focus();
        RefreshUI();
      }
    }

    private void OnCaptureNorthClicked(object sender, RoutedEventArgs e) => CaptureOrientation(SwitchOrientation.North);
    private void OnCaptureEastClicked(object sender, RoutedEventArgs e) => CaptureOrientation(SwitchOrientation.East);
    private void OnCaptureSouthClicked(object sender, RoutedEventArgs e) => CaptureOrientation(SwitchOrientation.South);
    private void OnCaptureWestClicked(object sender, RoutedEventArgs e) => CaptureOrientation(SwitchOrientation.West);

    private void OnAutoDeriveClicked(object sender, RoutedEventArgs e)
    {
      var typeConfig = _settings.GetTypeConfig(_currentType);
      if (!typeConfig.North.IsConfigured)
      {
        MessageBox.Show(
          $"Please capture the NORTH orientation for {typeConfig.DisplayName} first before auto-deriving East, South, and West.",
          "ACIES - Switch Configuration",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
        return;
      }

      SwitchConfigurationStore.AutoDeriveOrientations(
        typeConfig.North,
        SwitchOrientation.North,
        out var north,
        out var east,
        out var south,
        out var west);

      typeConfig.North = north;
      typeConfig.East = east;
      typeConfig.South = south;
      typeConfig.West = west;

      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc != null)
      {
        SwitchConfigurationStore.SaveSettings(doc.Database, _settings, doc);
      }

      RefreshUI();
      MessageBox.Show(
        $"Successfully auto-derived East, South, and West orientations from North for {typeConfig.DisplayName}!",
        "ACIES - Switch Configuration",
        MessageBoxButton.OK,
        MessageBoxImage.Information);
    }

    private void OnLoadDefaultsClicked(object sender, RoutedEventArgs e)
    {
      var loaded = SwitchConfigurationStore.LoadGlobalDefaults();
      if (loaded == null)
      {
        MessageBox.Show(
          $"No global default switch configurations found at:\n{SwitchConfigurationStore.GetGlobalDefaultsPath()}",
          "ACIES - Load Defaults",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
        return;
      }

      _settings = loaded;
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc != null)
      {
        SwitchConfigurationStore.SaveSettings(doc.Database, _settings, doc);
      }
      RefreshUI();
      MessageBox.Show("Global default switch configurations loaded successfully!", "ACIES", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnSaveDefaultsClicked(object sender, RoutedEventArgs e)
    {
      if (SwitchConfigurationStore.SaveGlobalDefaults(_settings, out string err))
      {
        MessageBox.Show(
          $"Current configurations saved as global default template at:\n{SwitchConfigurationStore.GetGlobalDefaultsPath()}",
          "ACIES - Save Defaults",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
      }
      else
      {
        MessageBox.Show($"Failed to save defaults: {err}", "ACIES", MessageBoxButton.OK, MessageBoxImage.Error);
      }
    }

    private void OnSaveProjectClicked(object sender, RoutedEventArgs e)
    {
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc != null)
      {
        if (SwitchConfigurationStore.SaveSettings(doc.Database, _settings, out string err, doc))
        {
          MessageBox.Show("Switch configurations saved project-wide!", "ACIES", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else
        {
          MessageBox.Show($"Error saving settings: {err}", "ACIES", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
      }
      else
      {
        MessageBox.Show("No active drawing open to save settings.", "ACIES", MessageBoxButton.OK, MessageBoxImage.Warning);
      }
    }

    private void OnCloseClicked(object sender, RoutedEventArgs e)
    {
      Close();
    }
  }
}
