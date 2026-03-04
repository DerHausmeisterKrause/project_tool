using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    private const double PeekWidth = 220;
    private const double ExpandedWidth = 456;
    private const double CollapsedHeight = 32;
    private const double ExpandedHeight = 300;
    private const double NotificationHeight = 126;
    private const double EdgeMargin = 10;

    private static int _instanceCounter;

    private DockAnchor _dockAnchor;
    private Vector _dockOffset = new(0, 0);
    private Rect _homeRect;

    private bool _isRightDragging;
    private Point _dragMouseOffset;
    private DateTime _lastDragMoveLogAt = DateTime.MinValue;
    private Storyboard? _stateStoryboard;
    private IslandVisualState _currentState = IslandVisualState.Collapsed;

    public DynamicIslandWindow()
    {
        InitializeComponent();
        DataContext = new DynamicIslandViewModel();
        if (DataContext is DynamicIslandViewModel vm)
            vm.PropertyChanged += OnViewModelPropertyChanged;

        _instanceCounter++;
        Log($"DynamicIslandWindow created. InstanceCount={_instanceCounter}");

        Loaded += (_, _) =>
        {
            _dockAnchor = LoadDockAnchor();
            ResetToHome(animate: false);
        };

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            var exStyle = GetWindowLong(hwnd, GwlExstyle);
            SetWindowLong(hwnd, GwlExstyle, exStyle | WsExToolwindow);
        };

        Deactivated += (_, _) =>
        {
            if (_isRightDragging)
                return;

            if (DataContext is DynamicIslandViewModel vm)
                vm.IsExpanded = false;

            ResetToHome();
        };

        Closed += (_, _) =>
        {
            if (DataContext is DynamicIslandViewModel vm)
            {
                vm.PropertyChanged -= OnViewModelPropertyChanged;
                vm.Stop();
            }

            _instanceCounter = Math.Max(0, _instanceCounter - 1);
            Log($"DynamicIslandWindow closed. InstanceCount={_instanceCounter}");
        };
    }

    public void EnqueueNotification(Guid taskId, string text, ReminderKind kind)
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return;

        vm.EnqueueNotification(taskId, text, kind);
        ResetToHome();
    }

    private void IslandRoot_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isRightDragging || DataContext is not DynamicIslandViewModel vm)
            return;

        vm.ToggleExpandCommand.Execute(null);
        ResetToHome();
        e.Handled = true;
    }

    private void Window_MouseEnter(object sender, MouseEventArgs e)
    {
        if (_isRightDragging)
            return;

        ResetToHome();
    }

    private void Window_MouseLeave(object sender, MouseEventArgs e)
    {
        if (_isRightDragging)
            return;

        ResetToHome();
    }

    private void Window_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isRightDragging = true;
        StopStateAnimation();
        _dragMouseOffset = e.GetPosition(this);
        CaptureMouse();
        Cursor = Cursors.SizeAll;
        Log("Drag start");
        e.Handled = true;
    }

    private void Window_MouseMove(object sender, MouseEventArgs e)
    {
        if (!_isRightDragging)
            return;

        var pos = e.GetPosition(null);
        Left = pos.X - _dragMouseOffset.X;
        Top = pos.Y - _dragMouseOffset.Y;

        if ((DateTime.Now - _lastDragMoveLogAt).TotalMilliseconds >= 160)
        {
            _lastDragMoveLogAt = DateTime.Now;
            Log($"Drag move Left={Left:F0} Top={Top:F0}");
        }
    }

    private void Window_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isRightDragging)
            return;

        _isRightDragging = false;
        ReleaseMouseCapture();
        Cursor = Cursors.Arrow;

        _dockAnchor = CalculateNearestDockAnchor();
        SaveDockAnchor(_dockAnchor);
        Log($"Drag end -> Snap {_dockAnchor}");

        ResetToHome();
        e.Handled = true;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(DynamicIslandViewModel.IsExpanded) or nameof(DynamicIslandViewModel.ActiveNotification))
            ResetToHome();
    }

    private IslandVisualState ResolveTargetState()
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return IslandVisualState.Collapsed;

        if (vm.HasNotification)
            return IslandVisualState.NotificationOverlay;

        if (vm.IsExpanded)
            return IslandVisualState.Expanded;

        return IsMouseOver ? IslandVisualState.Peek : IslandVisualState.Collapsed;
    }

    private void ResetToHome(bool animate = true)
    {
        ApplyVisualState(ResolveTargetState(), animate);
    }

    private void ApplyVisualState(IslandVisualState state, bool animate = true)
    {
        _homeRect = CalculateHomeRect(_dockAnchor, state, _dockOffset);

        if (!animate || _isRightDragging)
        {
            StopStateAnimation();
            Left = _homeRect.Left;
            Top = _homeRect.Top;
            Width = _homeRect.Width;
            Height = _homeRect.Height;
            UpdateState(state);
            return;
        }

        StopStateAnimation();
        var duration = TimeSpan.FromMilliseconds(170);

        _stateStoryboard = new Storyboard();
        AddAnimation(_stateStoryboard, this, LeftProperty, _homeRect.Left, duration);
        AddAnimation(_stateStoryboard, this, TopProperty, _homeRect.Top, duration);
        AddAnimation(_stateStoryboard, this, WidthProperty, _homeRect.Width, duration);
        AddAnimation(_stateStoryboard, this, HeightProperty, _homeRect.Height, duration);
        _stateStoryboard.Completed += (_, _) =>
        {
            Left = _homeRect.Left;
            Top = _homeRect.Top;
            Width = _homeRect.Width;
            Height = _homeRect.Height;
            UpdateState(state);
        };
        _stateStoryboard.Begin();
    }

    private static void AddAnimation(Storyboard sb, DependencyObject target, DependencyProperty property, double to, TimeSpan duration)
    {
        var anim = new DoubleAnimation
        {
            To = to,
            Duration = duration,
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        Storyboard.SetTarget(anim, target);
        Storyboard.SetTargetProperty(anim, new PropertyPath(property));
        sb.Children.Add(anim);
    }

    private void StopStateAnimation()
    {
        if (_stateStoryboard == null)
            return;

        _stateStoryboard.Stop();
        _stateStoryboard = null;
    }

    private void UpdateState(IslandVisualState newState)
    {
        if (_currentState == newState)
            return;

        _currentState = newState;
        Log($"State -> {newState}");
    }

    private static double WidthForState(IslandVisualState state) => state switch
    {
        IslandVisualState.Collapsed => CollapsedWidth,
        IslandVisualState.Peek => PeekWidth,
        _ => ExpandedWidth
    };

    private static double HeightForState(IslandVisualState state) => state switch
    {
        IslandVisualState.NotificationOverlay => NotificationHeight,
        IslandVisualState.Expanded => ExpandedHeight,
        _ => CollapsedHeight
    };

    private static Rect CalculateHomeRect(DockAnchor anchor, IslandVisualState state, Vector offset)
    {
        var area = SystemParameters.WorkArea;
        var width = WidthForState(state);
        var height = HeightForState(state);
        var isExpandedState = state is IslandVisualState.Expanded or IslandVisualState.NotificationOverlay;

        var centerX = area.Left + ((area.Width - width) / 2);
        var leftVisible = area.Left + EdgeMargin;
        var rightVisible = area.Right - width - EdgeMargin;
        var topVisible = area.Top + EdgeMargin;
        var topHalfHidden = area.Top - (height / 2);
        var bottomVisible = area.Bottom - height - EdgeMargin;
        var bottomHalfHidden = area.Bottom - (height / 2);

        (double left, double top) = anchor switch
        {
            DockAnchor.TopCenter => (centerX, isExpandedState ? topVisible : topHalfHidden),
            DockAnchor.TopLeft => (leftVisible, isExpandedState ? topVisible : topHalfHidden),
            DockAnchor.TopRight => (rightVisible, isExpandedState ? topVisible : topHalfHidden),
            DockAnchor.BottomLeft => (leftVisible, isExpandedState ? bottomVisible : bottomHalfHidden),
            DockAnchor.BottomCenter => (centerX, isExpandedState ? bottomVisible : bottomHalfHidden),
            DockAnchor.BottomRight => (rightVisible, isExpandedState ? bottomVisible : bottomHalfHidden),
            _ => (centerX, isExpandedState ? topVisible : topHalfHidden)
        };

        left += offset.X;
        top += offset.Y;
        return new Rect(left, top, width, height);
    }

    private DockAnchor CalculateNearestDockAnchor()
    {
        var area = SystemParameters.WorkArea;
        var centerX = Left + (Width / 2);
        var centerY = Top + (Height / 2);

        var points = new Dictionary<DockAnchor, Point>
        {
            [DockAnchor.TopCenter] = new(area.Left + area.Width / 2, area.Top),
            [DockAnchor.TopLeft] = new(area.Left, area.Top),
            [DockAnchor.TopRight] = new(area.Right, area.Top),
            [DockAnchor.BottomLeft] = new(area.Left, area.Bottom),
            [DockAnchor.BottomCenter] = new(area.Left + area.Width / 2, area.Bottom),
            [DockAnchor.BottomRight] = new(area.Right, area.Bottom)
        };

        return points.OrderBy(x => Distance(centerX, centerY, x.Value.X, x.Value.Y)).First().Key;
    }

    private static double Distance(double x1, double y1, double x2, double y2)
    {
        var dx = x1 - x2;
        var dy = y1 - y2;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    private DockAnchor LoadDockAnchor()
    {
        var raw = ServiceLocator.Settings.Current.DynamicIslandDockPosition?.Trim();
        return raw switch
        {
            nameof(DockAnchor.TopCenter) => DockAnchor.TopCenter,
            nameof(DockAnchor.TopLeft) => DockAnchor.TopLeft,
            nameof(DockAnchor.TopRight) => DockAnchor.TopRight,
            nameof(DockAnchor.BottomLeft) => DockAnchor.BottomLeft,
            nameof(DockAnchor.BottomCenter) => DockAnchor.BottomCenter,
            nameof(DockAnchor.BottomRight) => DockAnchor.BottomRight,
            // Legacy compatibility
            "LeftCenter" => DockAnchor.TopLeft,
            "RightCenter" => DockAnchor.TopRight,
            _ => DockAnchor.TopCenter
        };
    }

    private void SaveDockAnchor(DockAnchor anchor)
    {
        ServiceLocator.Settings.Current.DynamicIslandDockPosition = anchor.ToString();
        ServiceLocator.Settings.Save();
    }

    private static void Log(string message)
    {
        try { ServiceLocator.Logger.Info($"[DynamicIslandWindow] {message}"); } catch { }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}

public enum DockAnchor
{
    TopCenter,
    TopLeft,
    TopRight,
    BottomLeft,
    BottomCenter,
    BottomRight
}

public enum IslandVisualState
{
    Collapsed,
    Peek,
    Expanded,
    NotificationOverlay
}
