using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class DynamicIslandWindow : Window
{
    private const int GwlExstyle = -20;
    private const int WsExToolwindow = 0x00000080;

    private const double PeekWidth = 220;
    private const double ExpandedWidth = 456;
    private const double PeekHeight = 32;
    private const double ExpandedHeight = 300;
    private const double NotificationHeight = 126;
    private const double EdgeMargin = 10;
    private const int DragThrottleMs = 16;

    private static int _instanceCounter;

    private DockAnchor _dockAnchor;
    private Vector _dockOffset = new(0, 0);
    private Rect _homePeekRect;

    private IslandState _state = IslandState.Hidden;
    private Storyboard? _stateStoryboard;

    private bool _isDragging;
    private Point _dragOffsetInWindow;
    private DateTime _lastDragMoveAt = DateTime.MinValue;

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
            SetState(IslandState.Peek, "Loaded", animate: false);
        };

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            var exStyle = GetWindowLong(hwnd, GwlExstyle);
            SetWindowLong(hwnd, GwlExstyle, exStyle | WsExToolwindow);
        };

        Deactivated += (_, _) =>
        {
            if (_isDragging)
                return;

            if (DataContext is DynamicIslandViewModel vm)
                vm.IsExpanded = false;

            SetState(IslandState.Peek, "Window Deactivated");
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
        SetState(IslandState.NotificationOverlay, "Notification Enqueued");
    }

    private void IslandRoot_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isDragging || DataContext is not DynamicIslandViewModel vm)
            return;

        vm.ToggleExpandCommand.Execute(null);
        SetState(ResolveTargetStateFromVm(), "Left Click Toggle");
        e.Handled = true;
    }

    private void Window_MouseEnter(object sender, MouseEventArgs e)
    {
        if (_isDragging)
            return;

        if (ResolveTargetStateFromVm() == IslandState.Peek)
            SetState(IslandState.Expanded, "Mouse Enter Peek");
    }

    private void Window_MouseLeave(object sender, MouseEventArgs e)
    {
        if (_isDragging)
            return;

        if (DataContext is DynamicIslandViewModel vm && !vm.IsExpanded && !vm.HasNotification)
            SetState(IslandState.Peek, "Mouse Leave");
    }

    private void Window_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isDragging = true;
        _dragOffsetInWindow = e.GetPosition(this);
        CaptureMouse();
        Cursor = Cursors.SizeAll;

        SetState(IslandState.Dragging, "Drag Start", animate: false);
        e.Handled = true;
    }

    private void Window_MouseMove(object sender, MouseEventArgs e)
    {
        if (!_isDragging)
            return;

        var now = DateTime.UtcNow;
        if ((now - _lastDragMoveAt).TotalMilliseconds < DragThrottleMs)
            return;
        _lastDragMoveAt = now;

        var mouseScreen = GetMouseScreenDip();
        var newLeft = mouseScreen.X - _dragOffsetInWindow.X;
        var newTop = mouseScreen.Y - _dragOffsetInWindow.Y;

        Left = newLeft;
        Top = newTop;

        Log($"Drag move Left={Left:F0} Top={Top:F0}");
    }

    private void Window_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isDragging)
            return;

        _isDragging = false;
        ReleaseMouseCapture();
        Cursor = Cursors.Arrow;

        _dockAnchor = CalculateNearestDockAnchor();
        _dockOffset = new Vector(0, 0);
        SaveDockAnchor(_dockAnchor);

        Log($"Drag end -> Snap {_dockAnchor}");
        SetState(IslandState.Peek, "Drag End Snap");
        e.Handled = true;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(DynamicIslandViewModel.IsExpanded) or nameof(DynamicIslandViewModel.ActiveNotification))
            SetState(ResolveTargetStateFromVm(), $"VM Change: {e.PropertyName}");
    }

    private IslandState ResolveTargetStateFromVm()
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return IslandState.Hidden;

        if (vm.HasNotification)
            return IslandState.NotificationOverlay;

        if (vm.IsExpanded)
            return IslandState.Expanded;

        return IslandState.Peek;
    }

    private void SetState(IslandState newState, string reason, bool animate = true)
    {
        if (_isDragging && newState != IslandState.Dragging)
            return;

        StopStateAnimation();

        var rect = GetTargetRect(newState);
        var doAnimate = animate && newState != IslandState.Hidden && newState != IslandState.Dragging;

        if (!doAnimate)
        {
            ApplyRectHard(rect, newState);
            Log($"State -> {newState} ({reason})");
            return;
        }

        var duration = TimeSpan.FromMilliseconds(170);
        _stateStoryboard = new Storyboard();
        AddAnimation(_stateStoryboard, this, LeftProperty, rect.Left, duration);
        AddAnimation(_stateStoryboard, this, TopProperty, rect.Top, duration);
        AddAnimation(_stateStoryboard, this, WidthProperty, rect.Width, duration);
        AddAnimation(_stateStoryboard, this, HeightProperty, rect.Height, duration);
        _stateStoryboard.Completed += (_, _) =>
        {
            ApplyRectHard(rect, newState);
            Log($"State -> {newState} ({reason})");
        };
        _stateStoryboard.Begin(this, true);
    }

    private Rect GetTargetRect(IslandState state)
    {
        _homePeekRect = CalculateHomePeekRect(_dockAnchor, _dockOffset);

        return state switch
        {
            IslandState.Hidden => new Rect(_homePeekRect.Left, _homePeekRect.Top, _homePeekRect.Width, _homePeekRect.Height),
            IslandState.Peek => _homePeekRect,
            IslandState.Expanded => CalculateStateRect(_homePeekRect, ExpandedWidth, ExpandedHeight),
            IslandState.NotificationOverlay => CalculateStateRect(_homePeekRect, ExpandedWidth, NotificationHeight),
            IslandState.Dragging => new Rect(Left, Top, Width, Height),
            _ => _homePeekRect
        };
    }

    private static Rect CalculateStateRect(Rect basePeekRect, double targetWidth, double targetHeight)
    {
        var widthDelta = targetWidth - basePeekRect.Width;
        var heightDelta = targetHeight - basePeekRect.Height;
        var left = basePeekRect.Left - (widthDelta / 2d);
        var top = basePeekRect.Top - (heightDelta / 2d);
        return new Rect(left, top, targetWidth, targetHeight);
    }

    private static Rect CalculateHomePeekRect(DockAnchor anchor, Vector offset)
    {
        var area = SystemParameters.WorkArea;

        var centerX = area.Left + ((area.Width - PeekWidth) / 2);
        var leftVisible = area.Left + EdgeMargin;
        var rightVisible = area.Right - PeekWidth - EdgeMargin;
        var topHalfHidden = area.Top - (PeekHeight / 2);
        var bottomHalfHidden = area.Bottom - (PeekHeight / 2);

        (double left, double top) = anchor switch
        {
            DockAnchor.TopCenter => (centerX, topHalfHidden),
            DockAnchor.TopLeft => (leftVisible, topHalfHidden),
            DockAnchor.TopRight => (rightVisible, topHalfHidden),
            DockAnchor.BottomLeft => (leftVisible, bottomHalfHidden),
            DockAnchor.BottomCenter => (centerX, bottomHalfHidden),
            DockAnchor.BottomRight => (rightVisible, bottomHalfHidden),
            _ => (centerX, topHalfHidden)
        };

        left += offset.X;
        top += offset.Y;
        return new Rect(left, top, PeekWidth, PeekHeight);
    }

    private void ApplyRectHard(Rect rect, IslandState state)
    {
        Left = rect.Left;
        Top = rect.Top;
        Width = rect.Width;
        Height = rect.Height;

        if (state == IslandState.Hidden)
            Hide();
        else if (!IsVisible)
            Show();

        _state = state;
    }

    private void StopStateAnimation()
    {
        if (_stateStoryboard == null)
            return;

        _stateStoryboard.Stop(this);
        _stateStoryboard.Remove(this);
        _stateStoryboard = null;
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

    private Point GetMouseScreenDip()
    {
        var screenPx = PointToScreen(Mouse.GetPosition(this));
        var source = PresentationSource.FromVisual(this);
        if (source?.CompositionTarget == null)
            return screenPx;

        return source.CompositionTarget.TransformFromDevice.Transform(screenPx);
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

public enum IslandState
{
    Hidden,
    Peek,
    Expanded,
    NotificationOverlay,
    Dragging
}
