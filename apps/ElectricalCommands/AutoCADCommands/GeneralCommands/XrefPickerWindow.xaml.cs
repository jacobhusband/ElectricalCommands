using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace ElectricalCommands
{
    public class XrefFileItem
    {
        public string FullPath { get; set; }
        public string DisplayName { get; set; }
    }

    public partial class XrefPickerWindow : Window
    {
        public string SelectedFile { get; private set; }

        public XrefPickerWindow(string[] filePaths)
        {
            InitializeComponent();

            var items = filePaths
                .Select(f => new XrefFileItem
                {
                    FullPath = f,
                    DisplayName = Path.GetFileName(f)
                })
                .ToList();

            FileListBox.ItemsSource = items;

            if (items.Count > 0)
                FileListBox.SelectedIndex = 0;
        }

        private void ConfirmSelection()
        {
            var selected = FileListBox.SelectedItem as XrefFileItem;
            if (selected == null)
            {
                MessageBox.Show("Please select a file.", "No Selection",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedFile = selected.FullPath;
            DialogResult = true;
            Close();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            ConfirmSelection();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void FileListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (FileListBox.SelectedItem != null)
                ConfirmSelection();
        }
    }
}
