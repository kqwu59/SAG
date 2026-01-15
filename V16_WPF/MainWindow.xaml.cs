using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace NettoieXLSX.V16
{
    public partial class MainWindow : Window
    {
        private static readonly string[] ExcelExtensions = { ".xlsx" };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnFileDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                e.Effects = paths.Any(IsExcelFile) ? DragDropEffects.Copy : DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void OnFileDrop(object sender, DragEventArgs e)
        {
            if (sender is not TextBox textBox)
            {
                return;
            }

            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            var filePath = paths.FirstOrDefault(IsExcelFile);
            if (string.IsNullOrWhiteSpace(filePath))
            {
                MessageBox.Show(
                    "Merci de déposer un fichier .xlsx.",
                    "Format non pris en charge",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            textBox.Text = filePath;
        }

        private void OnBrowseCommandes(object sender, RoutedEventArgs e)
            => BrowseForExcelFile(CommandesTextBox);

        private void OnBrowseConstatations(object sender, RoutedEventArgs e)
            => BrowseForExcelFile(ConstatationsTextBox);

        private void OnBrowseFactures(object sender, RoutedEventArgs e)
            => BrowseForExcelFile(FacturesTextBox);

        private void OnBrowseEnvoiBdc(object sender, RoutedEventArgs e)
            => BrowseForExcelFile(EnvoiBdcTextBox);

        private void OnBrowseWorkflow(object sender, RoutedEventArgs e)
            => BrowseForExcelFile(WorkflowTextBox);

        private void OnBrowseOutput(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = "export_clean.xlsx",
            };

            if (dialog.ShowDialog(this) == true)
            {
                OutputTextBox.Text = dialog.FileName;
            }
        }

        private static void BrowseForExcelFile(TextBox target)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                CheckFileExists = true,
                Multiselect = false,
            };

            if (dialog.ShowDialog() == true)
            {
                target.Text = dialog.FileName;
            }
        }

        private static bool IsExcelFile(string path)
            => !string.IsNullOrWhiteSpace(path)
               && File.Exists(path)
               && ExcelExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);
    }
}
