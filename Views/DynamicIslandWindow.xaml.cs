using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls.Primitives;
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
    private const double ExpandedHeightNormal = 286;
    private const double ExpandedHeightNotificationMin = 108;
    private const double EdgeMargin = 10;
    private const double SafeVisibleMargin = 12;
    private const int DragThrottleMs = 16;

    private static int _instanceCounter;

    private DockAnchor _dockAnchor;
    private Vector _dockOffset = new(0, 0);

    private Storyboard? _stateStoryboard;
    private bool _isTransitionActive;
    private InteractionState _state = InteractionState.Collapsed;
    private InteractionState? _queuedStableState;

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
            _dockAnchor = DockAnchor.TopCenter;
            SetState(InteractionState.Collapsed, "Loaded TopCenter");
        };

        PreviewKeyDown += (_, evt) =>
        {
            if (evt.Key == Key.Escape)
            {
                SetState(InteractionState.Collapsed, "Esc Close");
                if (DataContext is DynamicIslandViewModel vm)
                    vm.IsExpanded = false;
                evt.Handled = true;
            }
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

            if (_state is InteractionState.Expanded or InteractionState.AnimatingOpen)
            {
                if (DataContext is DynamicIslandViewModel vm)
                    vm.IsExpanded = false;
                SetState(InteractionState.Collapsed, "Window Deactivated");
            }
        };

        LostMouseCapture += (_, _) => ReleaseDragCaptureIfNeeded();

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
        SetState(InteractionState.Expanded, "Notification Enqueued");
    }

    private void IslandRoot_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isDragging || DataContext is not DynamicIslandViewModel vm)
            return;

        if (IsFromButton(e.OriginalSource as DependencyObject))
            return;

        if (_state is InteractionState.Collapsed or InteractionState.AnimatingClose)
        {
            vm.IsExpanded = true;
            SetState(InteractionState.Expanded, "Left Click Open");
        }
        else if (_state is InteractionState.Expanded && !vm.HasNotification)
        {
            vm.IsExpanded = false;
            SetState(InteractionState.Collapsed, "Left Click Close");
        }

        e.Handled = true;
    }

    private void NotificationOpenButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return;

        vm.OpenNotificationCommand.Execute(null);
        if (!vm.HasNotification)
        {
            vm.IsExpanded = false;
            SetState(InteractionState.Collapsed, "Notification Open Button");
        }
        e.Handled = true;
    }

    private void NotificationCloseButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return;

        vm.DismissNotificationCommand.Execute(null);
        if (!vm.HasNotification)
        {
            vm.IsExpanded = false;
            SetState(InteractionState.Collapsed, "Notification Close Button");
        }
        else
        {
            SetState(InteractionState.Expanded, "Notification Next Item");
        }

        e.Handled = true;
    }

    private void Window_MouseEnter(object sender, MouseEventArgs e)
    {
        if (_isDragging)
            return;

        IslandRoot.Opacity = 1.0;
    }

    private void Window_MouseLeave(object sender, MouseEventArgs e)
    {
        if (_isDragging)
            return;

        IslandRoot.Opacity = 0.96;
    }

    private void Window_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isDragging = true;
        _dragOffsetInWindow = e.GetPosition(this);
        CaptureMouse();
        Cursor = Cursors.SizeAll;
        StopStateAnimation();
        _state = InteractionState.Collapsed;
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
        Left = mouseScreen.X - _dragOffsetInWindow.X;
        Top = mouseScreen.Y - _dragOffsetInWindow.Y;
    }

    private void Window_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isDragging)
            return;

        ReleaseDragCaptureIfNeeded();
        _dockAnchor = CalculateNearestDockAnchor();
        _dockOffset = new Vector(0, 0);
        SaveDockAnchor(_dockAnchor);

        if (DataContext is DynamicIslandViewModel vm)
            vm.IsExpanded = false;
        SetState(InteractionState.Collapsed, "Drag End Snap");
        e.Handled = true;
    }

    private void ReleaseDragCaptureIfNeeded()
    {
        if (!_isDragging)
            return;

        _isDragging = false;
        if (IsMouseCaptured)
            ReleaseMouseCapture();
        Cursor = Cursors.Arrow;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is not (nameof(DynamicIslandViewModel.IsExpanded) or nameof(DynamicIslandViewModel.ActiveNotification)))
            return;

        if (DataContext is not DynamicIslandViewModel vm)
            return;

        var target = vm.IsExpanded || vm.HasNotification
            ? InteractionState.Expanded
            : InteractionState.Collapsed;
        SetState(target, $"VM Change: {e.PropertyName}");
    }

    private static bool IsFromButton(DependencyObject? source)
    {
        while (source != null)
        {
            if (source is ButtonBase)
                return true;
            source = VisualTreeHelper.GetParent(source);
        }

        return false;
    }

    // Root cause: sporadische Open/Close-Fehler wurden durch konkurrierende Trigger verursacht
    // (Window/VM/Event-Race + parallele Animationen). Diese Methode ist der einzige State-Writer.
    private void SetState(InteractionState requestedState, string reason)
    {
        if (_isDragging)
            return;

        if (_isTransitionActive)
        {
            if (requestedState is InteractionState.Collapsed or InteractionState.Expanded)
                _queuedStableState = requestedState;
            return;
        }

        if (requestedState == InteractionState.AnimatingOpen || requestedState == InteractionState.AnimatingClose)
            return;

        var targetState = requestedState;
        if (_state == targetState)
            return;

        StopStateAnimation();
        _isTransitionActive = true;

        var startRect = GetCurrentOrFallbackRect();
        ApplyRect(startRect);

        var willOpen = targetState == InteractionState.Expanded;
        _state = willOpen ? InteractionState.AnimatingOpen : InteractionState.AnimatingClose;
        Log($"State -> {_state} ({reason})");

        var targetRect = willOpen ? GetExpandedRect() : GetPeekRect();
        var duration = TimeSpan.FromMilliseconds(170);

        _stateStoryboard = new Storyboard();
        AddAnimation(_stateStoryboard, this, LeftProperty, targetRect.Left, duration);
        AddAnimation(_stateStoryboard, this, TopProperty, targetRect.Top, duration);
        AddAnimation(_stateStoryboard, this, WidthProperty, targetRect.Width, duration);
        AddAnimation(_stateStoryboard, this, HeightProperty, targetRect.Height, duration);
        _stateStoryboard.Completed += (_, _) =>
        {
            StopStateAnimation();
            ApplyRect(targetRect);
            ApplyHostHeights(targetState);
            _state = targetState;
            _isTransitionActive = false;
            Log($"State -> {_state} ({reason})");

            if (_queuedStableState.HasValue)
            {
                var queued = _queuedStableState.Value;
                _queuedStableState = null;
                SetState(queued, "Queued state");
            }
        };
        _stateStoryboard.Begin(this, true);
    }

    private Rect GetCurrentOrFallbackRect()
        => Width > 1 && Height > 1 ? new Rect(Left, Top, Width, Height) : GetPeekRect();

    private Rect GetPeekRect() => CalculatePeekRect(_dockAnchor, _dockOffset);

    private Rect GetExpandedRect()
    {
        var hasNotification = (DataContext as DynamicIslandViewModel)?.HasNotification == true;
        var targetHeight = hasNotification ? GetNotificationExpandedHeight() : ExpandedHeightNormal;
        return CalculateVisibleRect(_dockAnchor, ExpandedWidth, targetHeight);
    }

    private double GetNotificationExpandedHeight()
    {
        if ((DataContext as DynamicIslandViewModel)?.HasNotification != true)
            return ExpandedHeightNotificationMin;

        UpdateLayout();
        var padding = IslandRoot.Padding.Top + IslandRoot.Padding.Bottom;
        var availableWidth = Math.Max(120, ExpandedWidth - IslandRoot.Padding.Left - IslandRoot.Padding.Right);
        NotificationOverlay.Measure(new Size(availableWidth, double.PositiveInfinity));
        var desired = NotificationOverlay.DesiredSize.Height;

        var target = Math.Ceiling(desired + padding);
        return Math.Max(ExpandedHeightNotificationMin, target);
    }

    private Rect CalculateVisibleRect(DockAnchor anchor, double targetWidth, double targetHeight)
    {
        var area = GetCurrentWorkingAreaDip();

        var centerLeft = area.Left + ((area.Width - targetWidth) / 2d);
        var leftEdge = area.Left + EdgeMargin;
        var rightEdge = area.Right - targetWidth - EdgeMargin;
        var topEdge = area.Top + SafeVisibleMargin;
        var bottomEdge = area.Bottom - targetHeight - SafeVisibleMargin;

        var left = anchor switch
        {
            DockAnchor.TopLeft or DockAnchor.BottomLeft => leftEdge,
            DockAnchor.TopRight or DockAnchor.BottomRight => rightEdge,
            _ => centerLeft
        };

        var top = anchor switch
        {
            DockAnchor.BottomLeft or DockAnchor.BottomCenter or DockAnchor.BottomRight => bottomEdge,
            _ => topEdge
        };

        left = Math.Max(area.Left + EdgeMargin, Math.Min(left, area.Right - targetWidth - EdgeMargin));
        top = Math.Max(area.Top + SafeVisibleMargin, Math.Min(top, area.Bottom - targetHeight - SafeVisibleMargin));

        return new Rect(left, top, targetWidth, targetHeight);
    }

    private Rect CalculatePeekRect(DockAnchor anchor, Vector offset)
    {
        var area = GetCurrentWorkingAreaDip();

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

    private void ApplyRect(Rect rect)
    {
        Left = rect.Left;
        Top = rect.Top;
        Width = rect.Width;
        Height = rect.Height;

        if (!IsVisible)
            Show();
    }

    private void ApplyHostHeights(InteractionState state)
    {
        var hasNotification = (DataContext as DynamicIslandViewModel)?.HasNotification == true;
        if (state == InteractionState.Expanded && hasNotification)
        {
            var nHeight = GetNotificationExpandedHeight();
            ContentHost.MinHeight = nHeight;
            NotificationOverlay.MinHeight = 0;
            ExpandedContentHost.MinHeight = 0;
            return;
        }

        if (state == InteractionState.Expanded)
        {
            ContentHost.MinHeight = ExpandedHeightNormal;
            ExpandedContentHost.MinHeight = 120;
            NotificationOverlay.MinHeight = 0;
            return;
        }

        ContentHost.MinHeight = 0;
        NotificationOverlay.MinHeight = 0;
        ExpandedContentHost.MinHeight = 0;
    }

    private Rect GetCurrentWorkingAreaDip()
    {
        IntPtr monitor = IntPtr.Zero;

        if (GetCursorPos(out var cursor))
            monitor = MonitorFromPoint(cursor, MonitorDefaulttonearest);

        if (monitor == IntPtr.Zero)
        {
            var main = Application.Current?.MainWindow;
            if (main != null)
            {
                var hwnd = new WindowInteropHelper(main).Handle;
                if (hwnd != IntPtr.Zero)
                    monitor = MonitorFromWindow(hwnd, MonitorDefaulttonearest);
            }
        }

        if (monitor == IntPtr.Zero)
        {
            var primaryPoint = new NativePoint { X = 0, Y = 0 };
            monitor = MonitorFromPoint(primaryPoint, MonitorDefaulttoprimary);
        }

        if (monitor == IntPtr.Zero)
            return SystemParameters.WorkArea;

        var info = new MonitorInfoEx { CbSize = Marshal.SizeOf<MonitorInfoEx>() };
        if (!GetMonitorInfo(monitor, ref info))
            return SystemParameters.WorkArea;

        var source = PresentationSource.FromVisual(this);
        var fromDevice = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
        var topLeft = fromDevice.Transform(new Point(info.RcWork.Left, info.RcWork.Top));
        var bottomRight = fromDevice.Transform(new Point(info.RcWork.Right, info.RcWork.Bottom));
        return new Rect(topLeft, bottomRight);
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
#if DEBUG
        try { ServiceLocator.Logger.Info($"[DynamicIslandWindow] {message}"); } catch { }
#endif
    }

    private const int MonitorDefaulttoprimary = 1;
    private const int MonitorDefaulttonearest = 2;

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct MonitorInfoEx
    {
        public int CbSize;
        public NativeRect RcMonitor;
        public NativeRect RcWork;
        public int DwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string SzDevice;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr MonitorFromPoint(NativePoint pt, int dwFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, int dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MonitorInfoEx lpmi);

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

public enum InteractionState
{
    Collapsed,
    Expanded,
    AnimatingOpen,
    AnimatingClose
}
