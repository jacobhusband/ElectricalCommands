using System;
using System.Windows;

namespace ElectricalCommands
{
  public partial class ProjectChecklistWindow : Window
  {
    private readonly IProjectChecklistEditorControl _editorControl;
    private bool _flushedPendingSave;

    internal ProjectChecklistWindow(IProjectChecklistEditorControl editorControl)
    {
      if (editorControl == null)
      {
        throw new ArgumentNullException(nameof(editorControl));
      }

      InitializeComponent();
      _editorControl = editorControl;
      ChecklistHost.Child = editorControl.HostControl;
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
      FlushPendingSaveOnce();
      ChecklistHost.Child = null;
      base.OnClosing(e);
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
      Close();
    }

    private void FlushPendingSaveOnce()
    {
      if (_flushedPendingSave)
      {
        return;
      }

      _flushedPendingSave = true;
      _editorControl.FlushPendingSave();
    }
  }
}
