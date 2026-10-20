using System.Windows;
using System.Windows.Controls;
using TaskTool.ViewModels;

namespace TaskTool.Views.Settings;

public partial class DemoFeaturesSettingsView : UserControl
{
    public DemoFeaturesSettingsView()
    {
        InitializeComponent();
        Loaded += (_, _) => RefreshState();
        DataContextChanged += (_, _) => RefreshState();
    }

    private void Unlock_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is SettingsViewModel viewModel)
            viewModel.TryUnlockDemoFeatures(MasterPasswordBox.Password);
        MasterPasswordBox.Clear();
        RefreshState();
    }

    private void Lock_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is SettingsViewModel viewModel)
            viewModel.LockDemoFeatures();
        MasterPasswordBox.Clear();
        RefreshState();
    }

    private void RefreshState()
    {
        var unlocked = DataContext is SettingsViewModel viewModel && viewModel.IsDemoFeaturesUnlocked;
        LockedPanel.Visibility = unlocked ? Visibility.Collapsed : Visibility.Visible;
        UnlockedPanel.Visibility = unlocked ? Visibility.Visible : Visibility.Collapsed;
    }
}
