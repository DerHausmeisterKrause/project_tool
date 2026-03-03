using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Input;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class DynamicIslandWindow : Window
{
    private const int GwlExstyle = -20;
    private const int WsExToolwindow = 0x00000080;

    public DynamicIslandWindow()
    {
        InitializeComponent();
        DataContext = new DynamicIslandViewModel();

        Loaded += (_, _) =>
        {
            PositionTopCenter();
            ApplySize();
        };

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            var exStyle = GetWindowLong(hwnd, GwlExstyle);
            SetWindowLong(hwnd, GwlExstyle, exStyle | WsExToolwindow);
        };

        Deactivated += (_, _) =>
        {
            if (DataContext is DynamicIslandViewModel vm)
            {
                vm.IsExpanded = false;
                ApplySize();
            }
        };
    }

    private void IslandRoot_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is not DynamicIslandViewModel vm) return;
        vm.ToggleExpandCommand.Execute(null);
        ApplySize();
        e.Handled = true;
    }

    private void ApplySize()
    {
        if (DataContext is not DynamicIslandViewModel vm) return;
        Width = vm.IsExpanded ? 430 : 220;
        Height = vm.IsExpanded ? 290 : 56;
        PositionTopCenter();
    }

    private void PositionTopCenter()
    {
        Left = SystemParameters.WorkArea.Left + (SystemParameters.WorkArea.Width - Width) / 2;
        Top = SystemParameters.WorkArea.Top + 10;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}
