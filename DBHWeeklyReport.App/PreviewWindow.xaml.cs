using System.Windows;
using DBHWeeklyReport.App.ViewModels;

namespace DBHWeeklyReport.App;

public partial class PreviewWindow : Window
{
    public PreviewWindow(PreviewWindowViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }

    public bool Confirmed { get; private set; }

    private void OnWriteClick(object sender, RoutedEventArgs e)
    {
        Confirmed = true;
        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        Confirmed = false;
        DialogResult = false;
        Close();
    }
}
