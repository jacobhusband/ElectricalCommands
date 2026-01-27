using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;

namespace ElectricalCommands
{
  /// <summary>
  /// View model for signature list items.
  /// </summary>
  public class SignatureViewModel
  {
    public string Id { get; set; }
    public string Name { get; set; }
    public string FilePath { get; set; }
    public bool IsDefault { get; set; }
    public bool IsSelected { get; set; }
  }

  /// <summary>
  /// Interaction logic for SignatureManagerWindow.xaml
  /// </summary>
  public partial class SignatureManagerWindow : Window
  {
    private T24SignatureSettings _settings;
    private List<SignatureViewModel> _viewModels;

    /// <summary>
    /// Gets whether the user confirmed their selection.
    /// </summary>
    public bool Confirmed { get; private set; }

    /// <summary>
    /// Gets the selected signature ID after the dialog closes.
    /// </summary>
    public string SelectedSignatureId { get; private set; }

    public SignatureManagerWindow(T24SignatureSettings settings)
    {
      InitializeComponent();

      _settings = settings ?? SignatureManager.LoadSettings();
      LoadSignatures();
    }

    private void LoadSignatures()
    {
      _viewModels = _settings.Signatures.Select(s => new SignatureViewModel
      {
        Id = s.Id,
        Name = s.Name,
        FilePath = s.FilePath,
        IsDefault = s.IsDefault,
        IsSelected = s.Id == _settings.SelectedSignatureId
      }).ToList();

      SignatureListBox.ItemsSource = _viewModels;

      // Select the current signature
      var selectedVm = _viewModels.FirstOrDefault(vm => vm.IsSelected);
      if (selectedVm != null)
      {
        SignatureListBox.SelectedItem = selectedVm;
      }

      UpdateStatus();
    }

    private void UpdateStatus()
    {
      var selected = SignatureManager.GetSelectedSignature(_settings);
      if (selected != null)
      {
        StatusText.Text = $"Active: {selected.Name}";
      }
      else
      {
        StatusText.Text = "Active: None";
      }

      // Update delete button state
      var selectedVm = SignatureListBox.SelectedItem as SignatureViewModel;
      DeleteButton.IsEnabled = selectedVm != null && !selectedVm.IsDefault;
    }

    private void UpdatePreview()
    {
      var selectedVm = SignatureListBox.SelectedItem as SignatureViewModel;

      if (selectedVm != null && File.Exists(selectedVm.FilePath))
      {
        try
        {
          var bitmap = new BitmapImage();
          bitmap.BeginInit();
          bitmap.CacheOption = BitmapCacheOption.OnLoad;
          bitmap.UriSource = new Uri(selectedVm.FilePath);
          bitmap.EndInit();
          bitmap.Freeze();

          PreviewImage.Source = bitmap;
        }
        catch
        {
          PreviewImage.Source = null;
        }
      }
      else
      {
        PreviewImage.Source = null;
      }
    }

    private void SignatureListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
      UpdatePreview();

      // Update delete button state
      var selectedVm = SignatureListBox.SelectedItem as SignatureViewModel;
      DeleteButton.IsEnabled = selectedVm != null && !selectedVm.IsDefault;
    }

    private void AddButton_Click(object sender, RoutedEventArgs e)
    {
      var dialog = new Microsoft.Win32.OpenFileDialog
      {
        Filter = "Image Files (*.png;*.gif;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff)|*.png;*.gif;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff",
        Title = "Select Signature Image"
      };

      if (dialog.ShowDialog() == true)
      {
        // Prompt for name
        string defaultName = Path.GetFileNameWithoutExtension(dialog.FileName);
        string name = PromptForName("Add Signature", "Enter a name for this signature:", defaultName);

        if (!string.IsNullOrWhiteSpace(name))
        {
          var entry = SignatureManager.AddSignature(_settings, dialog.FileName, name);
          if (entry != null)
          {
            SignatureManager.SaveSettings(_settings);
            LoadSignatures();

            // Select the new entry
            var newVm = _viewModels.FirstOrDefault(vm => vm.Id == entry.Id);
            if (newVm != null)
            {
              SignatureListBox.SelectedItem = newVm;
            }
          }
          else
          {
            MessageBox.Show("Failed to add signature. Please try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
          }
        }
      }
    }

    private void RenameButton_Click(object sender, RoutedEventArgs e)
    {
      var selectedVm = SignatureListBox.SelectedItem as SignatureViewModel;
      if (selectedVm == null)
        return;

      string newName = PromptForName("Rename Signature", "Enter a new name:", selectedVm.Name);

      if (!string.IsNullOrWhiteSpace(newName) && newName != selectedVm.Name)
      {
        if (SignatureManager.RenameSignature(_settings, selectedVm.Id, newName))
        {
          SignatureManager.SaveSettings(_settings);
          LoadSignatures();

          // Re-select
          var renamedVm = _viewModels.FirstOrDefault(vm => vm.Id == selectedVm.Id);
          if (renamedVm != null)
          {
            SignatureListBox.SelectedItem = renamedVm;
          }
        }
      }
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
      var selectedVm = SignatureListBox.SelectedItem as SignatureViewModel;
      if (selectedVm == null || selectedVm.IsDefault)
        return;

      var result = MessageBox.Show(
        $"Are you sure you want to delete '{selectedVm.Name}'?",
        "Confirm Delete",
        MessageBoxButton.YesNo,
        MessageBoxImage.Question);

      if (result == MessageBoxResult.Yes)
      {
        if (SignatureManager.RemoveSignature(_settings, selectedVm.Id))
        {
          SignatureManager.SaveSettings(_settings);
          LoadSignatures();
        }
      }
    }

    private void UseSelectedButton_Click(object sender, RoutedEventArgs e)
    {
      var selectedVm = SignatureListBox.SelectedItem as SignatureViewModel;
      if (selectedVm != null)
      {
        _settings.SelectedSignatureId = selectedVm.Id;
        SignatureManager.SaveSettings(_settings);

        SelectedSignatureId = selectedVm.Id;
        Confirmed = true;
        DialogResult = true;
        Close();
      }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
      Confirmed = false;
      DialogResult = false;
      Close();
    }

    /// <summary>
    /// Shows a simple input dialog to prompt for a name.
    /// </summary>
    private string PromptForName(string title, string prompt, string defaultValue)
    {
      // Create a simple input dialog
      var inputWindow = new Window
      {
        Title = title,
        Width = 350,
        Height = 150,
        WindowStartupLocation = WindowStartupLocation.CenterOwner,
        Owner = this,
        ResizeMode = ResizeMode.NoResize,
        ShowInTaskbar = false
      };

      var grid = new System.Windows.Controls.Grid { Margin = new Thickness(12) };
      grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
      grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
      grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });

      var label = new System.Windows.Controls.TextBlock { Text = prompt, Margin = new Thickness(0, 0, 0, 8) };
      System.Windows.Controls.Grid.SetRow(label, 0);

      var textBox = new System.Windows.Controls.TextBox { Text = defaultValue, Margin = new Thickness(0, 0, 0, 12) };
      System.Windows.Controls.Grid.SetRow(textBox, 1);

      var buttonPanel = new System.Windows.Controls.StackPanel
      {
        Orientation = System.Windows.Controls.Orientation.Horizontal,
        HorizontalAlignment = HorizontalAlignment.Right
      };
      System.Windows.Controls.Grid.SetRow(buttonPanel, 2);

      var okButton = new System.Windows.Controls.Button { Content = "OK", Width = 75, IsDefault = true };
      var cancelButton = new System.Windows.Controls.Button { Content = "Cancel", Width = 75, Margin = new Thickness(8, 0, 0, 0), IsCancel = true };

      string result = null;
      okButton.Click += (s, args) =>
      {
        result = textBox.Text?.Trim();
        inputWindow.DialogResult = true;
        inputWindow.Close();
      };
      cancelButton.Click += (s, args) =>
      {
        inputWindow.DialogResult = false;
        inputWindow.Close();
      };

      buttonPanel.Children.Add(okButton);
      buttonPanel.Children.Add(cancelButton);

      grid.Children.Add(label);
      grid.Children.Add(textBox);
      grid.Children.Add(buttonPanel);

      inputWindow.Content = grid;

      // Select all text on load
      inputWindow.Loaded += (s, args) =>
      {
        textBox.SelectAll();
        textBox.Focus();
      };

      if (inputWindow.ShowDialog() == true)
      {
        return result;
      }

      return null;
    }
  }
}
