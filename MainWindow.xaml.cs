using System.Windows;
using TaskTool.Services;

namespace TaskTool;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = ServiceLocator.MainViewModel;
    }
}
