using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ElectricalCommands
{
  internal static class HomerunSettingsPalette
  {
    private static readonly Guid PaletteGuid = new Guid("F3CBAC46-1A30-4EB0-971F-4A33904A0D6B");
    private static PaletteSet _palette;
    private static HomerunSettingsPaletteControl _control;
    private static bool _documentEventsAttached;

    internal static void Show()
    {
      EnsureCreated();
      Refresh();
      _palette.Visible = true;
    }

    internal static void Refresh()
    {
      if (_control == null)
      {
        return;
      }

      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        _control.LoadSnapshot(HomerunSettingsSnapshot.Empty());
        _control.SetEnabled(false);
        _control.SetStatus("Open a drawing to configure home runs.");
        return;
      }

      try
      {
        _control.LoadSnapshot(ReadSnapshot(document.Database));
        _control.SetEnabled(true);
      }
      catch (System.Exception ex)
      {
        _control.SetEnabled(false);
        _control.SetStatus("Unable to read drawing settings: " + ex.Message);
      }
    }

    internal static void SetStatus(string message)
    {
      _control?.SetStatus(message);
    }

    private static void EnsureCreated()
    {
      if (_palette != null)
      {
        return;
      }

      _control = new HomerunSettingsPaletteControl();
      _palette = new PaletteSet("Home Run Settings", PaletteGuid)
      {
        Style = PaletteSetStyles.ShowAutoHideButton |
          PaletteSetStyles.ShowCloseButton |
          PaletteSetStyles.ShowPropertiesMenu,
        DockEnabled = DockSides.Left | DockSides.Right,
        MinimumSize = new Size(360, 340),
      };
      _palette.Add("Settings", _control);

      if (!_documentEventsAttached)
      {
        AcApplication.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
        _documentEventsAttached = true;
      }
    }

    private static void DocumentManager_DocumentActivated(
      object sender,
      DocumentCollectionEventArgs e
    )
    {
      if (_palette?.Visible == true)
      {
        Refresh();
      }
    }

    private static HomerunSettingsSnapshot ReadSnapshot(Database database)
    {
      var snapshot = new HomerunSettingsSnapshot();

      if (ElectricalDrawingSettingsStore.TryReadPanelLocation(database, out var location))
      {
        snapshot.LocationText = string.Format(
          CultureInfo.InvariantCulture,
          "{0:0.####}, {1:0.####}, {2:0.####} ({3})",
          location.Point.X,
          location.Point.Y,
          location.Point.Z,
          location.Context
        );
      }

      if (ElectricalDrawingSettingsStore.TryReadScale(database, out var scale))
      {
        snapshot.ScaleText = scale.DisplayText;
      }

      if (ElectricalDrawingSettingsStore.TryReadPanelName(database, out string panelName))
      {
        snapshot.PanelName = panelName;
      }

      GeneralCommands.ResolveHomerunLayerId(database, out string selectedLayerName);
      snapshot.SelectedLayerName = selectedLayerName;

      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        LayerTable layers = transaction.GetObject(
          database.LayerTableId,
          OpenMode.ForRead
        ) as LayerTable;
        if (layers != null)
        {
          foreach (ObjectId id in layers)
          {
            if (id.IsErased)
            {
              continue;
            }

            LayerTableRecord layer = transaction.GetObject(
              id,
              OpenMode.ForRead
            ) as LayerTableRecord;
            if (layer != null)
            {
              snapshot.LayerNames.Add(layer.Name);
            }
          }
        }
      }

      snapshot.LayerNames = snapshot.LayerNames
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
        .ToList();
      return snapshot;
    }
  }

  internal sealed class HomerunSettingsPaletteControl : UserControl
  {
    private static readonly string[] StandardScales =
    {
      "1\" = 1'-0\"",
      "3/4\" = 1'-0\"",
      "1/2\" = 1'-0\"",
      "3/8\" = 1'-0\"",
      "1/4\" = 1'-0\"",
      "3/16\" = 1'-0\"",
      "1/8\" = 1'-0\"",
      "3/32\" = 1'-0\"",
      "1/16\" = 1'-0\"",
    };

    private readonly Label _locationValue;
    private readonly Button _pickLocationButton;
    private readonly ComboBox _scaleComboBox;
    private readonly Button _setScaleButton;
    private readonly TextBox _panelNameTextBox;
    private readonly Button _setPanelNameButton;
    private readonly ComboBox _layerComboBox;
    private readonly Button _setLayerButton;
    private readonly Button _runHomerunButton;
    private readonly Button _refreshButton;
    private readonly Label _statusLabel;

    internal HomerunSettingsPaletteControl()
    {
      AutoScaleMode = AutoScaleMode.Font;
      Padding = new Padding(12);

      var layout = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        ColumnCount = 3,
        RowCount = 8,
      };
      layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

      Label heading = new Label
      {
        AutoSize = true,
        Font = new System.Drawing.Font(Font, FontStyle.Bold),
        Margin = new Padding(0, 0, 0, 12),
        Text = "Home Run Drawing Settings",
      };
      layout.Controls.Add(heading, 0, 0);
      layout.SetColumnSpan(heading, 3);

      _locationValue = new Label
      {
        AutoEllipsis = true,
        Dock = DockStyle.Fill,
        Margin = new Padding(0, 6, 8, 10),
        Text = "Not set",
        TextAlign = ContentAlignment.MiddleLeft,
      };
      _pickLocationButton = CreateActionButton("Pick Location", PickLocationButton_Click);
      AddRow(layout, 1, "Location", _locationValue, _pickLocationButton);

      _scaleComboBox = new ComboBox
      {
        Dock = DockStyle.Top,
        DropDownStyle = ComboBoxStyle.DropDown,
        Margin = new Padding(0, 3, 8, 10),
      };
      _scaleComboBox.Items.AddRange(StandardScales);
      _setScaleButton = CreateActionButton("Set Scale", SetScaleButton_Click);
      AddRow(layout, 2, "Scale", _scaleComboBox, _setScaleButton);

      _panelNameTextBox = new TextBox
      {
        Dock = DockStyle.Top,
        Margin = new Padding(0, 3, 8, 10),
      };
      _setPanelNameButton = CreateActionButton("Set Name", SetPanelNameButton_Click);
      AddRow(layout, 3, "Panel", _panelNameTextBox, _setPanelNameButton);

      _layerComboBox = new ComboBox
      {
        Dock = DockStyle.Top,
        DropDownStyle = ComboBoxStyle.DropDownList,
        Margin = new Padding(0, 3, 8, 10),
        Sorted = true,
      };
      _setLayerButton = CreateActionButton("Set Layer", SetLayerButton_Click);
      AddRow(layout, 4, "HR Layer", _layerComboBox, _setLayerButton);

      _runHomerunButton = new Button
      {
        AutoSize = true,
        Dock = DockStyle.Top,
        Font = new System.Drawing.Font(Font, FontStyle.Bold),
        Margin = new Padding(0, 8, 0, 8),
        Text = "Run Home Run (HR)",
      };
      _runHomerunButton.Click += RunHomerunButton_Click;
      layout.Controls.Add(_runHomerunButton, 0, 5);
      layout.SetColumnSpan(_runHomerunButton, 3);

      _refreshButton = new Button
      {
        AutoSize = true,
        Anchor = AnchorStyles.Left,
        Margin = new Padding(0, 4, 8, 4),
        Text = "Refresh",
      };
      _refreshButton.Click += (_, __) =>
      {
        HomerunSettingsPalette.Refresh();
        SetStatus("Settings refreshed from the current drawing.");
      };
      layout.Controls.Add(_refreshButton, 0, 6);

      _statusLabel = new Label
      {
        AutoSize = true,
        Dock = DockStyle.Fill,
        ForeColor = SystemColors.GrayText,
        Margin = new Padding(0, 8, 0, 0),
        Text = "Ready.",
      };
      layout.Controls.Add(_statusLabel, 0, 7);
      layout.SetColumnSpan(_statusLabel, 3);

      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      Controls.Add(layout);
    }

    internal void LoadSnapshot(HomerunSettingsSnapshot snapshot)
    {
      snapshot = snapshot ?? HomerunSettingsSnapshot.Empty();
      _locationValue.Text = string.IsNullOrWhiteSpace(snapshot.LocationText)
        ? "Not set"
        : snapshot.LocationText;
      _scaleComboBox.Text = snapshot.ScaleText ?? string.Empty;
      _panelNameTextBox.Text = snapshot.PanelName ?? string.Empty;

      _layerComboBox.Items.Clear();
      foreach (string layerName in snapshot.LayerNames)
      {
        _layerComboBox.Items.Add(layerName);
      }

      int selectedIndex = FindComboItem(_layerComboBox, snapshot.SelectedLayerName);
      if (selectedIndex < 0 && _layerComboBox.Items.Count > 0)
      {
        selectedIndex = 0;
      }
      _layerComboBox.SelectedIndex = selectedIndex;
    }

    internal void SetEnabled(bool enabled)
    {
      _pickLocationButton.Enabled = enabled;
      _scaleComboBox.Enabled = enabled;
      _setScaleButton.Enabled = enabled;
      _panelNameTextBox.Enabled = enabled;
      _setPanelNameButton.Enabled = enabled;
      _layerComboBox.Enabled = enabled;
      _setLayerButton.Enabled = enabled;
      _runHomerunButton.Enabled = enabled;
      _refreshButton.Enabled = enabled;
    }

    internal void SetStatus(string message)
    {
      _statusLabel.Text = string.IsNullOrWhiteSpace(message) ? "Ready." : message;
    }

    private static Button CreateActionButton(string text, EventHandler clickHandler)
    {
      var button = new Button
      {
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 10),
        Text = text,
      };
      button.Click += clickHandler;
      return button;
    }

    private static void AddRow(
      TableLayoutPanel layout,
      int row,
      string labelText,
      Control input,
      Control button
    )
    {
      var label = new Label
      {
        AutoSize = true,
        Margin = new Padding(0, 7, 10, 10),
        Text = labelText,
      };
      layout.Controls.Add(label, 0, row);
      layout.Controls.Add(input, 1, row);
      layout.Controls.Add(button, 2, row);
    }

    private static int FindComboItem(ComboBox comboBox, string value)
    {
      for (int i = 0; i < comboBox.Items.Count; i++)
      {
        if (string.Equals(
          Convert.ToString(comboBox.Items[i]),
          value,
          StringComparison.OrdinalIgnoreCase
        ))
        {
          return i;
        }
      }
      return -1;
    }

    private void PickLocationButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the panel location.");
        return;
      }

      SetStatus("Pick the panel location in the drawing.");
      document.SendStringToExecute("_.SPL ", true, false, false);
    }

    private void SetScaleButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the scale.");
        return;
      }

      string candidate = (_scaleComboBox.Text ?? string.Empty).Trim();
      if (!GeneralCommands.TryParseDrawingScale(
        candidate,
        out double paperInchesPerFoot,
        out string displayText
      ))
      {
        SetStatus("Enter a scale such as 1/4, 3/16, or 1:48.");
        return;
      }

      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          ElectricalDrawingSettingsStore.WriteScale(
            document.Database,
            paperInchesPerFoot,
            displayText
          );
        }
        HomerunSettingsPalette.Refresh();
        SetStatus("Scale set to " + displayText + ".");
        document.Editor.WriteMessage("\nHR scale set to " + displayText + ".");
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to set the scale: " + ex.Message);
      }
    }

    private void SetPanelNameButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the panel name.");
        return;
      }

      string panelName = (_panelNameTextBox.Text ?? string.Empty).Trim();
      if (panelName.Length == 0)
      {
        SetStatus("Panel name cannot be blank.");
        return;
      }

      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          ElectricalDrawingSettingsStore.WritePanelName(document.Database, panelName);
        }
        HomerunSettingsPalette.Refresh();
        SetStatus("Panel name set to " + panelName + ".");
        document.Editor.WriteMessage("\nHR panel name set to " + panelName + ".");
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to set the panel name: " + ex.Message);
      }
    }

    private void SetLayerButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the HR layer.");
        return;
      }

      string layerName = Convert.ToString(_layerComboBox.SelectedItem) ?? string.Empty;
      if (layerName.Length == 0)
      {
        SetStatus("Select an existing layer.");
        return;
      }

      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          using (Transaction transaction = document.Database.TransactionManager.StartOpenCloseTransaction())
          {
            LayerTable layers = transaction.GetObject(
              document.Database.LayerTableId,
              OpenMode.ForRead
            ) as LayerTable;
            if (layers == null || !layers.Has(layerName))
            {
              throw new InvalidOperationException("The selected layer no longer exists.");
            }
          }

          ElectricalDrawingSettingsStore.WriteHomerunLayer(document.Database, layerName);
        }
        HomerunSettingsPalette.Refresh();
        SetStatus("HR layer set to " + layerName + ".");
        document.Editor.WriteMessage("\nHR layer set to " + layerName + ".");
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to set the HR layer: " + ex.Message);
      }
    }

    private void RunHomerunButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before running HR.");
        return;
      }

      SetStatus("Starting HR in the current drawing.");
      document.SendStringToExecute("_.HR ", true, false, false);
    }
  }

  internal sealed class HomerunSettingsSnapshot
  {
    internal string LocationText { get; set; } = string.Empty;
    internal string ScaleText { get; set; } = string.Empty;
    internal string PanelName { get; set; } = string.Empty;
    internal string SelectedLayerName { get; set; } = string.Empty;
    internal List<string> LayerNames { get; set; } = new List<string>();

    internal static HomerunSettingsSnapshot Empty()
    {
      return new HomerunSettingsSnapshot();
    }
  }
}
