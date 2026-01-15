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
        private const string IntroText =
            "La présente macro permet d’avoir une vision globale du traitement des commandes de la base 0018.\n\n" +
            "Pour ce faire, elle exploite plusieurs fichiers au format EXCEL :\n" +
            "- 3 extractions sous Geslab :\n" +
            "   « commandes / réservations » : RAJOUTER dans les paramètres d’affichage : Type de flux et Auteur\n" +
            "   « constatation »\n" +
            "   « facture »\n" +
            "- l’extraction des « workflows » sous DMF\n" +
            "- le fichier « Envoi BDC » sous SAG/TUTOS complété lors du traitement des bons de commande\n\n" +
            "Pour garantir une bonne utilisation de la macro, mettez les bons fichiers sur la bonne ligne correspondante\n" +
            "(Astuce : glissez vos fichiers .xlsx sur les champs !)\n\n" +
            "Dans les fichiers extraits de Geslab, seules les lignes sous 'Liste des résultats' seront prises en compte.\n";

        public MainWindow()
        {
            InitializeComponent();
            WireEvents();
            LogTextBox.Text = IntroText;
        }

        private void WireEvents()
        {
            RegisterDropHandlers(CommandesTextBox);
            RegisterDropHandlers(ConstatationsTextBox);
            RegisterDropHandlers(FacturesTextBox);
            RegisterDropHandlers(EnvoiBdcTextBox);
            RegisterDropHandlers(WorkflowTextBox);

            BrowseCommandesButton.Click += OnBrowseCommandes;
            BrowseConstatationsButton.Click += OnBrowseConstatations;
            BrowseFacturesButton.Click += OnBrowseFactures;
            BrowseEnvoiBdcButton.Click += OnBrowseEnvoiBdc;
            BrowseWorkflowButton.Click += OnBrowseWorkflow;
            BrowseOutputButton.Click += OnBrowseOutput;
            RunButton.Click += OnRun;
            ClearButton.Click += OnClear;
        }

        private void RegisterDropHandlers(TextBox target)
        {
            target.PreviewDragOver += OnFileDragOver;
            target.Drop += OnFileDrop;
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

        private void OnRun(object sender, RoutedEventArgs e)
        {
            var outputPath = OutputTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                MessageBox.Show("Veuillez choisir un fichier de sortie (.xlsx).",
                    "Sortie manquante",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                outputPath += ".xlsx";
                OutputTextBox.Text = outputPath;
            }

            RunButton.IsEnabled = false;
            StatusTextBlock.Text = "Traitement en cours…";
            LogTextBox.AppendText("\nDémarrage du traitement...\n");
            try
            {
                ExcelProcessor.Process(
                    CommandesTextBox.Text,
                    ConstatationsTextBox.Text,
                    FacturesTextBox.Text,
                    EnvoiBdcTextBox.Text,
                    WorkflowTextBox.Text,
                    outputPath,
                    msg => LogTextBox.AppendText(msg + "\n"));
                StatusTextBlock.Text = "Terminé";
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Erreur";
                LogTextBox.AppendText($"Erreur: {ex.Message}\n");
                MessageBox.Show(ex.ToString(), "Erreur de traitement", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                RunButton.IsEnabled = true;
            }
        }

        private void OnClear(object sender, RoutedEventArgs e)
        {
            CommandesTextBox.Text = string.Empty;
            ConstatationsTextBox.Text = string.Empty;
            FacturesTextBox.Text = string.Empty;
            EnvoiBdcTextBox.Text = string.Empty;
            WorkflowTextBox.Text = string.Empty;
            OutputTextBox.Text = string.Empty;
            LogTextBox.Text = IntroText;
            StatusTextBlock.Text = "Prêt";
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
