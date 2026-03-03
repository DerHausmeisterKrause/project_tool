using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using TaskTool.Services;

namespace TaskTool;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = ServiceLocator.MainViewModel;
        TryLoadWindowIcon();
    }

    private void TryLoadWindowIcon()
    {
        try
        {
            Icon = new BitmapImage(new Uri("pack://application:,,,/Assets/Plenaro.ico", UriKind.Absolute));
            return;
        }
        catch
        {
            // fallback to loose file next to executable
        }

        var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "Plenaro.ico");
        if (!File.Exists(iconPath))
            return;

        Icon = new BitmapImage(new Uri(iconPath, UriKind.Absolute));
    }
}
