using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace DocxViewer;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void OpenButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Word documents (*.docx)|*.docx",
            CheckFileExists = true,
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            LoadDocument(dialog.FileName);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Could not open document:\n{ex.Message}", "Docx Viewer",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadDocument(string path)
    {
        // FileShare.ReadWrite (+ Delete) means we never take an exclusive lock: other users/apps,
        // including Word itself, can still open, edit and save the same file while we are viewing it.
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);

        var document = DocxRenderer.Render(stream);
        Viewer.Document = document;
        StatusText.Text = $"Loaded (read-only, shared): {path}";
    }
}
