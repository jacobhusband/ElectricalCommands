using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Drawing;
using System.Windows.Forms;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ElectricalCommands
{
  internal static class SwitchCircuitPalette
  {
    private static readonly Guid PaletteGuid = new Guid("D1446284-2CF0-4CC4-B7A0-AC38E2F731D4");
    private static PaletteSet _palette;
    private static SwitchCircuitPaletteControl _control;

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
        _control.SetDrawingState(string.Empty, false);
        _control.SetStatus("Open a drawing to assign switch circuits.");
        return;
      }

      try
      {
        int nextIndex = SwitchCircuitCommands.GetEffectiveNextIndex(document.Database);
        _control.SetDrawingState(SwitchCircuitSequence.ToSuffix(nextIndex), true);
      }
      catch (System.Exception ex)
      {
        _control.SetDrawingState(string.Empty, false);
        _control.SetStatus("Unable to read drawing state: " + ex.Message);
      }
    }

    internal static void SetStatus(string message)
    {
      _control?.SetStatus(message);
    }

    internal static void Dispose()
    {
      if (_palette != null)
      {
        _palette.Dispose();
        _palette = null;
      }

      _control = null;
    }

    private static void EnsureCreated()
    {
      if (_palette != null)
      {
        return;
      }

      _control = new SwitchCircuitPaletteControl();
      _palette = new PaletteSet("Switch Circuit", PaletteGuid)
      {
        Style = PaletteSetStyles.ShowAutoHideButton |
          PaletteSetStyles.ShowCloseButton |
          PaletteSetStyles.ShowPropertiesMenu,
        DockEnabled = DockSides.Left | DockSides.Right,
        MinimumSize = new Size(260, 190),
      };
      _palette.Add("Circuiting", _control);
    }
  }

  internal sealed class SwitchCircuitPaletteControl : UserControl
  {
    private readonly TextBox _nextSuffixTextBox;
    private readonly Button _setNextButton;
    private readonly Button _applyButton;
    private readonly Label _statusLabel;

    internal SwitchCircuitPaletteControl()
    {
      AutoScaleMode = AutoScaleMode.Font;
      Padding = new Padding(12);

      TableLayoutPanel layout = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        AutoSize = true,
        ColumnCount = 2,
        RowCount = 4,
      };
      layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

      Label heading = new Label
      {
        Text = "Next control circuit",
        AutoSize = true,
        Font = new System.Drawing.Font(Font, FontStyle.Bold),
        Margin = new Padding(0, 0, 0, 6),
      };
      layout.Controls.Add(heading, 0, 0);
      layout.SetColumnSpan(heading, 2);

      _nextSuffixTextBox = new TextBox
      {
        Dock = DockStyle.Top,
        CharacterCasing = CharacterCasing.Lower,
        Margin = new Padding(0, 0, 8, 10),
      };
      _setNextButton = new Button
      {
        Text = "Set Next",
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 10),
      };
      _setNextButton.Click += SetNextButton_Click;
      layout.Controls.Add(_nextSuffixTextBox, 0, 1);
      layout.Controls.Add(_setNextButton, 1, 1);

      _applyButton = new Button
      {
        Text = "Apply to MText Selection",
        Dock = DockStyle.Top,
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 10),
      };
      _applyButton.Click += ApplyButton_Click;
      layout.Controls.Add(_applyButton, 0, 2);
      layout.SetColumnSpan(_applyButton, 2);

      _statusLabel = new Label
      {
        Text = "Select fixture-tag MText, then click Apply.",
        Dock = DockStyle.Fill,
        AutoSize = true,
        ForeColor = SystemColors.GrayText,
      };
      layout.Controls.Add(_statusLabel, 0, 3);
      layout.SetColumnSpan(_statusLabel, 2);

      Controls.Add(layout);
    }

    internal void SetDrawingState(string nextSuffix, bool enabled)
    {
      _nextSuffixTextBox.Text = nextSuffix ?? string.Empty;
      _nextSuffixTextBox.Enabled = enabled;
      _setNextButton.Enabled = enabled;
      _applyButton.Enabled = enabled;
    }

    internal void SetStatus(string message)
    {
      _statusLabel.Text = message ?? string.Empty;
    }

    private void ApplyButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing to assign switch circuits.");
        return;
      }

      SetStatus("Select fixture-tag MText in the drawing.");
      document.SendStringToExecute("-SWITCHCIRCUITAPPLY ", true, false, false);
    }

    private void SetNextButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before changing the counter.");
        return;
      }

      string candidate = (_nextSuffixTextBox.Text ?? string.Empty).Trim();
      if (!SwitchCircuitSequence.TryParseSuffix(candidate, out int requestedIndex))
      {
        SetStatus("Enter lowercase letters only; i, l, and o are not allowed.");
        return;
      }

      SetStatus("Updating the drawing counter...");
      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          if (!SwitchCircuitCommands.TrySetNextIndex(
            document.Database,
            requestedIndex,
            out string message
          ))
          {
            SetStatus(message);
            SwitchCircuitPalette.Refresh();
            return;
          }

          SwitchCircuitPalette.Refresh();
          SetStatus(message);
          document.Editor.WriteMessage("\n" + message);
        }
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to update the switch circuit counter: " + ex.Message);
        SwitchCircuitPalette.Refresh();
      }
    }
  }
}
