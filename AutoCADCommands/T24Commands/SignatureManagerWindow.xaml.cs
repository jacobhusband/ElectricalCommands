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
    public bool IsActive { get; set; }
  }

  /// <summary>
  /// View model for template list items.
  /// </summary>
  public class TemplateViewModel
  {
    public string Id { get; set; }
    public string Name { get; set; }
    public bool IsDefault { get; set; }
    public bool IsActive { get; set; }
  }

  /// <summary>
  /// Interaction logic for SignatureManagerWindow.xaml
  /// </summary>
  public partial class SignatureManagerWindow : Window
  {
    private T24SignatureSettings _settings;
    private List<SignatureViewModel> _signatureViewModels;
    private List<TemplateViewModel> _templateViewModels;
    private bool _isLoadingTemplate;
    private bool _isLoadingSignature;

    /// <summary>
    /// Gets whether the user confirmed their selection.
    /// </summary>
    public bool Confirmed { get; private set; }

    /// <summary>
    /// Gets the selected template ID after the dialog closes.
    /// </summary>
    public string SelectedTemplateId { get; private set; }

    /// <summary>
    /// Gets the selected signature ID after the dialog closes.
    /// </summary>
    public string SelectedSignatureId { get; private set; }

    public SignatureManagerWindow(T24SignatureSettings settings)
    {
      InitializeComponent();

      _settings = settings ?? SignatureManager.LoadSettings();
      LoadSignatures();
      LoadTemplates();
      UpdateStatus();
    }

    private void LoadTemplates()
    {
      _isLoadingTemplate = true;

      _templateViewModels = _settings.Templates.Select(t => new TemplateViewModel
      {
        Id = t.Id,
        Name = t.Name,
        IsDefault = t.IsDefault,
        IsActive = t.Id == _settings.SelectedTemplateId
      }).ToList();

      TemplateListBox.ItemsSource = _templateViewModels;

      var selectedVm = _templateViewModels.FirstOrDefault(vm => vm.Id == _settings.SelectedTemplateId) ??
                       _templateViewModels.FirstOrDefault();
      if (selectedVm != null)
      {
        TemplateListBox.SelectedItem = selectedVm;
      }

      _isLoadingTemplate = false;
      LoadTemplateDetails(GetSelectedTemplate());
    }

    private void LoadSignatures()
    {
      _isLoadingSignature = true;

      _signatureViewModels = _settings.Signatures.Select(s => new SignatureViewModel
      {
        Id = s.Id,
        Name = s.Name,
        FilePath = s.FilePath,
        IsDefault = s.IsDefault,
        IsActive = s.Id == _settings.SelectedSignatureId
      }).ToList();

      SignatureListBox.ItemsSource = _signatureViewModels;

      var selectedVm = _signatureViewModels.FirstOrDefault(vm => vm.Id == _settings.SelectedSignatureId) ??
                       _signatureViewModels.FirstOrDefault();
      if (selectedVm != null)
      {
        SignatureListBox.SelectedItem = selectedVm;
      }

      TemplateSignatureComboBox.ItemsSource = _settings.Signatures;
      bool wasLoadingTemplate = _isLoadingTemplate;
      _isLoadingTemplate = true;
      UpdateTemplateSignatureSelection();
      _isLoadingTemplate = wasLoadingTemplate;

      _isLoadingSignature = false;
      UpdateSignaturePreview();
    }

    private void UpdateStatus()
    {
      var activeTemplate = SignatureManager.GetSelectedTemplate(_settings);
      var activeSignature = SignatureManager.GetSignatureForTemplate(_settings, activeTemplate);

      if (activeTemplate != null)
      {
        string signatureName = activeSignature?.Name ?? "None";
        StatusText.Text = $"Active Template: {activeTemplate.Name} (Signature: {signatureName})";
      }
      else
      {
        StatusText.Text = "Active Template: None";
      }
    }

    private T24TemplateEntry GetSelectedTemplate()
    {
      var selectedVm = TemplateListBox.SelectedItem as TemplateViewModel;
      if (selectedVm == null)
        return null;

      return _settings.Templates.FirstOrDefault(t => t.Id == selectedVm.Id);
    }

    private SignatureViewModel GetSelectedSignatureViewModel()
    {
      return SignatureListBox.SelectedItem as SignatureViewModel;
    }

    private SignatureEntry GetSelectedSignatureEntry()
    {
      var selectedVm = GetSelectedSignatureViewModel();
      if (selectedVm == null)
        return null;

      return _settings.Signatures.FirstOrDefault(s => s.Id == selectedVm.Id);
    }

    private void LoadTemplateDetails(T24TemplateEntry template)
    {
      _isLoadingTemplate = true;

      bool hasTemplate = template != null;
      ResponsibleNameTextBox.IsEnabled = hasTemplate;
      CompanyTextBox.IsEnabled = hasTemplate;
      Address1TextBox.IsEnabled = hasTemplate;
      Address2TextBox.IsEnabled = hasTemplate;
      PhoneTextBox.IsEnabled = hasTemplate;
      LicenseTextBox.IsEnabled = hasTemplate;
      TemplateSignatureComboBox.IsEnabled = hasTemplate;
      ManageSignaturesButton.IsEnabled = true;

      if (template == null)
      {
        ResponsibleNameTextBox.Text = string.Empty;
        CompanyTextBox.Text = string.Empty;
        Address1TextBox.Text = string.Empty;
        Address2TextBox.Text = string.Empty;
        PhoneTextBox.Text = string.Empty;
        LicenseTextBox.Text = string.Empty;
        TemplateSignatureComboBox.SelectedItem = null;
        TemplatePreviewImage.Source = null;
        DeleteTemplateButton.IsEnabled = false;
        _isLoadingTemplate = false;
        return;
      }

      ResponsibleNameTextBox.Text = template.ResponsibleName ?? string.Empty;
      CompanyTextBox.Text = template.Company ?? string.Empty;
      Address1TextBox.Text = template.AddressLine1 ?? string.Empty;
      Address2TextBox.Text = template.AddressLine2 ?? string.Empty;
      PhoneTextBox.Text = template.Phone ?? string.Empty;
      LicenseTextBox.Text = template.License ?? string.Empty;

      UpdateTemplateSignatureSelection();
      UpdateTemplatePreview(template);

      DeleteTemplateButton.IsEnabled = !template.IsDefault;
      _isLoadingTemplate = false;
    }

    private void UpdateTemplateSignatureSelection()
    {
      var template = GetSelectedTemplate();
      if (template == null)
      {
        TemplateSignatureComboBox.SelectedItem = null;
        return;
      }

      var signature = _settings.Signatures.FirstOrDefault(s => s.Id == template.SignatureId);
      TemplateSignatureComboBox.SelectedItem = signature;
    }

    private void UpdateSignaturePreview()
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

      var selected = GetSelectedSignatureViewModel();
      DeleteSignatureButton.IsEnabled = selected != null && !selected.IsDefault;
    }

    private void UpdateTemplatePreview(T24TemplateEntry template)
    {
      var signature = SignatureManager.GetSignatureForTemplate(_settings, template);
      if (signature != null && File.Exists(signature.FilePath))
      {
        try
        {
          var bitmap = new BitmapImage();
          bitmap.BeginInit();
          bitmap.CacheOption = BitmapCacheOption.OnLoad;
          bitmap.UriSource = new Uri(signature.FilePath);
          bitmap.EndInit();
          bitmap.Freeze();

          TemplatePreviewImage.Source = bitmap;
        }
        catch
        {
          TemplatePreviewImage.Source = null;
        }
      }
      else
      {
        TemplatePreviewImage.Source = null;
      }
    }

    private void TemplateListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
      if (_isLoadingTemplate)
        return;

      var template = GetSelectedTemplate();
      LoadTemplateDetails(template);
    }

    private void TemplateField_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
      if (_isLoadingTemplate)
        return;

      var template = GetSelectedTemplate();
      if (template == null)
        return;

      template.ResponsibleName = ResponsibleNameTextBox.Text ?? string.Empty;
      template.Company = CompanyTextBox.Text ?? string.Empty;
      template.AddressLine1 = Address1TextBox.Text ?? string.Empty;
      template.AddressLine2 = Address2TextBox.Text ?? string.Empty;
      template.Phone = PhoneTextBox.Text ?? string.Empty;
      template.License = LicenseTextBox.Text ?? string.Empty;

      SignatureManager.SaveSettings(_settings);
      UpdateStatus();
    }

    private void TemplateSignatureComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
      if (_isLoadingTemplate)
        return;

      var template = GetSelectedTemplate();
      if (template == null)
        return;

      var signature = TemplateSignatureComboBox.SelectedItem as SignatureEntry;
      if (signature == null)
        return;

      template.SignatureId = signature.Id;
      _settings.SelectedSignatureId = signature.Id;
      SignatureManager.SaveSettings(_settings);

      UpdateTemplatePreview(template);
      _isLoadingTemplate = true;
      LoadSignatures();
      _isLoadingTemplate = false;
      UpdateStatus();
    }

    private void ManageSignaturesButton_Click(object sender, RoutedEventArgs e)
    {
      ManagerTabs.SelectedIndex = 1;
    }

    private void AddTemplateButton_Click(object sender, RoutedEventArgs e)
    {
      var selectedTemplate = GetSelectedTemplate();
      string name = PromptForName("Add Template", "Enter a template name:", selectedTemplate?.Name ?? "New Template");

      if (!string.IsNullOrWhiteSpace(name))
      {
        var entry = SignatureManager.AddTemplate(_settings, selectedTemplate, name);
        if (entry != null)
        {
          SignatureManager.SaveSettings(_settings);
          LoadTemplates();

          var newVm = _templateViewModels.FirstOrDefault(vm => vm.Id == entry.Id);
          if (newVm != null)
          {
            TemplateListBox.SelectedItem = newVm;
          }
        }
        else
        {
          MessageBox.Show("Failed to add template. Please try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
      }
    }

    private void RenameTemplateButton_Click(object sender, RoutedEventArgs e)
    {
      var selectedVm = TemplateListBox.SelectedItem as TemplateViewModel;
      if (selectedVm == null)
        return;

      string newName = PromptForName("Rename Template", "Enter a new template name:", selectedVm.Name);

      if (!string.IsNullOrWhiteSpace(newName) && newName != selectedVm.Name)
      {
        if (SignatureManager.RenameTemplate(_settings, selectedVm.Id, newName))
        {
          SignatureManager.SaveSettings(_settings);
          LoadTemplates();

          var renamedVm = _templateViewModels.FirstOrDefault(vm => vm.Id == selectedVm.Id);
          if (renamedVm != null)
          {
            TemplateListBox.SelectedItem = renamedVm;
          }
          UpdateStatus();
        }
      }
    }

    private void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
    {
      var selectedVm = TemplateListBox.SelectedItem as TemplateViewModel;
      if (selectedVm == null || selectedVm.IsDefault)
        return;

      var result = MessageBox.Show(
        $"Are you sure you want to delete '{selectedVm.Name}'?",
        "Confirm Delete",
        MessageBoxButton.YesNo,
        MessageBoxImage.Question);

      if (result == MessageBoxResult.Yes)
      {
        if (SignatureManager.RemoveTemplate(_settings, selectedVm.Id))
        {
          SignatureManager.SaveSettings(_settings);
          LoadTemplates();
          UpdateStatus();
        }
      }
    }

    private void SignatureListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
      if (_isLoadingSignature)
        return;

      UpdateSignaturePreview();
    }

    private void AddSignatureButton_Click(object sender, RoutedEventArgs e)
    {
      var dialog = new Microsoft.Win32.OpenFileDialog
      {
        Filter = "Image Files (*.png;*.gif;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff)|*.png;*.gif;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff",
        Title = "Select Signature Image"
      };

      if (dialog.ShowDialog() == true)
      {
        string defaultName = Path.GetFileNameWithoutExtension(dialog.FileName);
        string name = PromptForName("Add Signature", "Enter a name for this signature:", defaultName);

        if (!string.IsNullOrWhiteSpace(name))
        {
          var entry = SignatureManager.AddSignature(_settings, dialog.FileName, name);
          if (entry != null)
          {
            SignatureManager.SaveSettings(_settings);
            LoadSignatures();

            var newVm = _signatureViewModels.FirstOrDefault(vm => vm.Id == entry.Id);
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

    private void RenameSignatureButton_Click(object sender, RoutedEventArgs e)
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

          var renamedVm = _signatureViewModels.FirstOrDefault(vm => vm.Id == selectedVm.Id);
          if (renamedVm != null)
          {
            SignatureListBox.SelectedItem = renamedVm;
          }
          UpdateStatus();
        }
      }
    }

    private void DeleteSignatureButton_Click(object sender, RoutedEventArgs e)
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
          LoadTemplates();
          UpdateStatus();
        }
      }
    }

    private void UseSelectedButton_Click(object sender, RoutedEventArgs e)
    {
      var selectedVm = TemplateListBox.SelectedItem as TemplateViewModel;
      if (selectedVm != null)
      {
        _settings.SelectedTemplateId = selectedVm.Id;

        var template = _settings.Templates.FirstOrDefault(t => t.Id == selectedVm.Id);
        if (template != null)
        {
          _settings.SelectedSignatureId = template.SignatureId;
        }

        SignatureManager.SaveSettings(_settings);

        SelectedTemplateId = selectedVm.Id;
        SelectedSignatureId = template?.SignatureId;
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
