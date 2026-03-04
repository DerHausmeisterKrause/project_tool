using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Animation;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class DynamicIslandWindow : Window
{
    private const int GwlExstyle = -20;
    private const int WsExToolwindow = 0x00000080;

    private const double CollapsedWidth = 176;
    private const double ExpandedWidth = 456;
    private const double CollapsedHeight = 32;
    private const double ExpandedHeight = 300;

    private bool _hovering;
    private bool _isRightDragging;
    private Point _dragMouseOffset;

    public DynamicIslandWindow()
    {
        InitializeComponent();
        DataContext = new DynamicIslandViewModel();

        Loaded += (_, _) =>
        {
            ApplyWindowState(animate: false);
            RepositionToSavedDock();
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
                ApplyWindowState();
            }
        };

        Closed += (_, _) =>
        {
            if (DataContext is DynamicIslandViewModel vm)
                vm.Stop();
        };
    }

    public void EnqueueNotification(Guid taskId, string text, Services.ReminderKind kind)
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return;

        vm.EnqueueNotification(taskId, text, kind);
        ApplyWindowState();
    }

    private void IslandRoot_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is not DynamicIslandViewModel vm) return;
        vm.ToggleExpandCommand.Execute(null);
        ApplyWindowState();
        e.Handled = true;
    }

    private void Window_MouseEnter(object sender, MouseEventArgs e)
    {
        _hovering = true;
        ApplyWindowState();
    }

    private void Window_MouseLeave(object sender, MouseEventArgs e)
    {
        _hovering = false;
        ApplyWindowState();
    }

    private void Window_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isRightDragging = true;
        _dragMouseOffset = e.GetPosition(this);
        CaptureMouse();
        Cursor = Cursors.SizeAll;
        e.Handled = true;
    }

    private void Window_MouseMove(object sender, MouseEventArgs e)
    {
        if (!_isRightDragging)
            return;

        var pos = e.GetPosition(null);
        Left = pos.X - _dragMouseOffset.X;
        Top = pos.Y - _dragMouseOffset.Y;
    }

    private void Window_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isRightDragging)
            return;

        _isRightDragging = false;
        ReleaseMouseCapture();
        Cursor = Cursors.Arrow;

        var dock = CalculateNearestDockPosition();
        SaveDock(dock);
        ApplyDockPosition(dock);
        e.Handled = true;
    }

    private DynamicIslandDockPosition CalculateNearestDockPosition()
    {
        var area = SystemParameters.WorkArea;
        var centerX = Left + (Width / 2);
        var centerY = Top + (Height / 2);

        var points = new Dictionary<DynamicIslandDockPosition, Point>
        {
            [DynamicIslandDockPosition.TopCenter] = new(area.Left + area.Width / 2, area.Top),
            [DynamicIslandDockPosition.TopLeft] = new(area.Left, area.Top),
            [DynamicIslandDockPosition.TopRight] = new(area.Right, area.Top),
            [DynamicIslandDockPosition.LeftCenter] = new(area.Left, area.Top + area.Height / 2),
            [DynamicIslandDockPosition.RightCenter] = new(area.Right, area.Top + area.Height / 2),
            [DynamicIslandDockPosition.BottomLeft] = new(area.Left, area.Bottom),
            [DynamicIslandDockPosition.BottomCenter] = new(area.Left + area.Width / 2, area.Bottom),
            [DynamicIslandDockPosition.BottomRight] = new(area.Right, area.Bottom)
        };

        return points.OrderBy(x => Distance(centerX, centerY, x.Value.X, x.Value.Y)).First().Key;
    }

    private static double Distance(double x1, double y1, double x2, double y2)
    {
        var dx = x1 - x2;
        var dy = y1 - y2;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    private void ApplyWindowState(bool animate = true)
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return;

        var targetWidth = vm.IsExpanded ? ExpandedWidth : (_hovering ? 204 : CollapsedWidth);
        var targetHeight = vm.IsExpanded ? ExpandedHeight : CollapsedHeight;

        if (!animate)
        {
            Width = targetWidth;
            Height = targetHeight;
            RepositionToSavedDock();
            return;
        }

        var widthAnim = new DoubleAnimation(targetWidth, TimeSpan.FromMilliseconds(180));
        var heightAnim = new DoubleAnimation(targetHeight, TimeSpan.FromMilliseconds(180));
        widthAnim.Completed += (_, _) => RepositionToSavedDock();

        BeginAnimation(WidthProperty, widthAnim, HandoffBehavior.SnapshotAndReplace);
        BeginAnimation(HeightProperty, heightAnim, HandoffBehavior.SnapshotAndReplace);
    }

    private void RepositionToSavedDock()
    {
        var raw = ServiceLocator.Settings.Current.DynamicIslandDockPosition;
        if (!Enum.TryParse<DynamicIslandDockPosition>(raw, true, out var dock))
            dock = DynamicIslandDockPosition.TopCenter;

        ApplyDockPosition(dock);
    }

    private void SaveDock(DynamicIslandDockPosition dock)
    {
        ServiceLocator.Settings.Current.DynamicIslandDockPosition = dock.ToString();
        ServiceLocator.Settings.Save();
    }

    private void ApplyDockPosition(DynamicIslandDockPosition dock)
    {
        var area = SystemParameters.WorkArea;
        var margin = 10d;
        var halfHiddenOffset = -(Height / 2);

        (double left, double top) = dock switch
        {
            DynamicIslandDockPosition.TopCenter => (area.Left + ((area.Width - Width) / 2), area.Top + halfHiddenOffset),
            DynamicIslandDockPosition.TopLeft => (area.Left + margin, area.Top + halfHiddenOffset),
            DynamicIslandDockPosition.TopRight => (area.Right - Width - margin, area.Top + halfHiddenOffset),
            DynamicIslandDockPosition.LeftCenter => (area.Left - (Width / 2), area.Top + ((area.Height - Height) / 2)),
            DynamicIslandDockPosition.RightCenter => (area.Right - (Width / 2), area.Top + ((area.Height - Height) / 2)),
            DynamicIslandDockPosition.BottomLeft => (area.Left + margin, area.Bottom - (Height / 2)),
            DynamicIslandDockPosition.BottomCenter => (area.Left + ((area.Width - Width) / 2), area.Bottom - (Height / 2)),
            DynamicIslandDockPosition.BottomRight => (area.Right - Width - margin, area.Bottom - (Height / 2)),
            _ => (area.Left + ((area.Width - Width) / 2), area.Top + halfHiddenOffset)
        };

        Left = left;
        Top = top;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}
