using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace NettoieXLSX.V16;

public partial class MainWindow : Window
{
    private const string IntroLog =
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
        "Dans les fichiers extraits de Geslab, seules les lignes sous 'Liste des résultats' seront prises en compte.";

    public MainWindow()
    {
        InitializeComponent();
        LogTextBox.Text = IntroLog;
        SourceInitialized += (_, _) => EnableAcrylic();
    }

    private void BrowseCommandes_Click(object sender, RoutedEventArgs e)
    {
        CommandesTextBox.Text = OpenFileDialog();
    }

    private void BrowseConstatations_Click(object sender, RoutedEventArgs e)
    {
        ConstatationsTextBox.Text = OpenFileDialog();
    }

    private void BrowseFactures_Click(object sender, RoutedEventArgs e)
    {
        FacturesTextBox.Text = OpenFileDialog();
    }

    private void BrowseEnvoiBdc_Click(object sender, RoutedEventArgs e)
    {
        EnvoiBdcTextBox.Text = OpenFileDialog();
    }

    private void BrowseWorkflow_Click(object sender, RoutedEventArgs e)
    {
        WorkflowTextBox.Text = OpenFileDialog();
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx",
            FileName = "NettoieXLSX_Export.xlsx"
        };

        if (dialog.ShowDialog() == true)
        {
            OutputTextBox.Text = dialog.FileName;
        }
    }

    private async void RunButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(CommandesTextBox.Text))
        {
            MessageBox.Show("Le fichier Commandes est obligatoire.", "Fichier manquant");
            return;
        }

        if (string.IsNullOrWhiteSpace(OutputTextBox.Text))
        {
            MessageBox.Show("Le fichier de sortie est obligatoire.", "Fichier manquant");
            return;
        }

        RunButton.IsEnabled = false;
        StatusTextBlock.Text = "Traitement en cours...";
        AppendLog("Démarrage du traitement...");

        try
        {
            await Task.Run(() =>
                ExcelProcessor.Process(
                    CommandesTextBox.Text,
                    ConstatationsTextBox.Text,
                    FacturesTextBox.Text,
                    EnvoiBdcTextBox.Text,
                    WorkflowTextBox.Text,
                    OutputTextBox.Text,
                    AppendLog));

            StatusTextBlock.Text = "Terminé";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Erreur";
            MessageBox.Show($"Erreur pendant le traitement : {ex.Message}", "Erreur");
        }
        finally
        {
            RunButton.IsEnabled = true;
        }
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        CommandesTextBox.Text = string.Empty;
        ConstatationsTextBox.Text = string.Empty;
        FacturesTextBox.Text = string.Empty;
        EnvoiBdcTextBox.Text = string.Empty;
        WorkflowTextBox.Text = string.Empty;
        OutputTextBox.Text = string.Empty;
        StatusTextBlock.Text = "Prêt";
        LogTextBox.Text = IntroLog;
    }

    private void FileTextBox_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void FileTextBox_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        if (sender is not System.Windows.Controls.TextBox textBox) return;
        var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
        if (files is { Length: > 0 })
        {
            textBox.Text = files[0];
        }
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ButtonState == MouseButtonState.Pressed)
        {
            DragMove();
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private string OpenFileDialog()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx"
        };

        return dialog.ShowDialog() == true ? dialog.FileName : string.Empty;
    }

    private void AppendLog(string message)
    {
        Dispatcher.Invoke(() =>
        {
            LogTextBox.Text += $"\n{message}";
            LogTextBox.ScrollToEnd();
        });
    }

    private void EnableAcrylic()
    {
        var windowHelper = new System.Windows.Interop.WindowInteropHelper(this);
        var accent = new AccentPolicy
        {
            AccentState = AccentState.AccentEnableAcrylicBlurBehind,
            GradientColor = 0xCCF7FBFF
        };

        var accentStructSize = Marshal.SizeOf(accent);
        var accentPtr = Marshal.AllocHGlobal(accentStructSize);
        Marshal.StructureToPtr(accent, accentPtr, false);

        var data = new WindowCompositionAttributeData
        {
            Attribute = WindowCompositionAttribute.WcaAccentPolicy,
            SizeOfData = accentStructSize,
            Data = accentPtr
        };

        SetWindowCompositionAttribute(windowHelper.Handle, ref data);
        Marshal.FreeHGlobal(accentPtr);
    }

    [DllImport("user32.dll")]
    private static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

    private enum AccentState
    {
        AccentDisabled = 0,
        AccentEnableGradient = 1,
        AccentEnableTransparentGradient = 2,
        AccentEnableBlurBehind = 3,
        AccentEnableAcrylicBlurBehind = 4
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct AccentPolicy
    {
        public AccentState AccentState;
        public int AccentFlags;
        public int GradientColor;
        public int AnimationId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WindowCompositionAttributeData
    {
        public WindowCompositionAttribute Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }

    private enum WindowCompositionAttribute
    {
        WcaAccentPolicy = 19
    }
}
