using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
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
    private const double ExpandedHeightNormal = 300;
    private const double ExpandedHeightNotification = 170;
    private const double EdgeMargin = 10;
    private const double SafeVisibleMargin = 12;
    private const int DragThrottleMs = 16;

    private static int _instanceCounter;

    private DockAnchor _dockAnchor;
    private Vector _dockOffset = new(0, 0);
    private Rect _homePeekRect;

    private IslandState _state = IslandState.Hidden;
    private Storyboard? _stateStoryboard;
    private bool _isStateTransitionInProgress;
    private IslandState? _queuedState;
    private string _queuedReason = string.Empty;
    private bool _hasAppliedInitialState;

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
            SetState(IslandState.Peek, "Loaded");
        };

        PreviewKeyDown += (_, evt) =>
        {
            if (evt.Key == Key.Escape && _state == IslandState.Expanded && DataContext is DynamicIslandViewModel vm && !vm.HasNotification)
            {
                vm.IsExpanded = false;
                SetState(IslandState.Peek, "Esc Close");
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

            if (DataContext is DynamicIslandViewModel vm)
                vm.IsExpanded = false;

            if (_state is IslandState.Expanded or IslandState.NotificationOverlay)
                SetState(IslandState.Peek, "Window Deactivated");
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
        SetState(IslandState.NotificationOverlay, "Notification Enqueued");
    }

    private void IslandRoot_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isDragging || DataContext is not DynamicIslandViewModel vm)
            return;

        if (IsFromButton(e.OriginalSource as DependencyObject))
            return;

        if (_state == IslandState.Peek)
        {
            vm.IsExpanded = true;
            SetState(IslandState.Expanded, "Left Click Open");
        }
        else if (_state == IslandState.Expanded && !vm.HasNotification)
        {
            vm.IsExpanded = false;
            SetState(IslandState.Peek, "Left Click Close");
        }

        e.Handled = true;
    }

    private void NotificationOpenButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return;

        vm.OpenNotificationCommand.Execute(null);
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
            SetState(IslandState.Peek, "Notification Close Button");
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

        SetState(IslandState.Dragging, "Drag Start");
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

    }

    private void Window_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isDragging)
            return;

        ReleaseDragCaptureIfNeeded();

        _dockAnchor = CalculateNearestDockAnchor();
        _dockOffset = new Vector(0, 0);
        SaveDockAnchor(_dockAnchor);

        Log($"Drag end -> Snap {_dockAnchor}");
        SetState(IslandState.Peek, "Drag End Snap");
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
        if (e.PropertyName is nameof(DynamicIslandViewModel.IsExpanded) or nameof(DynamicIslandViewModel.ActiveNotification))
            SetState(ResolveTargetStateFromVm(), $"VM Change: {e.PropertyName}");
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

    private void SetState(IslandState newState, string reason)
    {
        if (_isStateTransitionInProgress)
        {
            _queuedState = newState;
            _queuedReason = reason;
            return;
        }

        if (_state == newState && _hasAppliedInitialState)
            return;

        _isStateTransitionInProgress = true;
        StopStateAnimation();

        var startRect = GetSafeCurrentRect();
        Left = startRect.Left;
        Top = startRect.Top;
        Width = startRect.Width;
        Height = startRect.Height;

        var targetRect = GetTargetRect(newState);
        var animate = ShouldAnimateTransition(newState);

        if (!animate)
        {
            CompleteStateTransition(newState, targetRect, reason);
            return;
        }

        var duration = TimeSpan.FromMilliseconds(170);
        _stateStoryboard = new Storyboard();
        AddAnimation(_stateStoryboard, this, LeftProperty, targetRect.Left, duration);
        AddAnimation(_stateStoryboard, this, TopProperty, targetRect.Top, duration);
        AddAnimation(_stateStoryboard, this, WidthProperty, targetRect.Width, duration);
        AddAnimation(_stateStoryboard, this, HeightProperty, targetRect.Height, duration);
        _stateStoryboard.Completed += (_, _) => CompleteStateTransition(newState, targetRect, reason);
        _stateStoryboard.Begin(this, true);
    }

    private bool ShouldAnimateTransition(IslandState newState)
    {
        if (!_hasAppliedInitialState)
            return false;

        if (_isDragging)
            return false;

        if (newState is IslandState.Hidden or IslandState.Dragging)
            return false;

        return true;
    }

    private Rect GetSafeCurrentRect()
    {
        if (Width > 1 && Height > 1)
            return new Rect(Left, Top, Width, Height);

        return CalculateHomePeekRect(_dockAnchor, _dockOffset);
    }

    private void CompleteStateTransition(IslandState state, Rect targetRect, string reason)
    {
        ApplyRectHard(targetRect, state);
        Log($"State -> {state} ({reason})");

        _hasAppliedInitialState = true;
        _isStateTransitionInProgress = false;

        if (_queuedState.HasValue)
        {
            var queuedState = _queuedState.Value;
            var queuedReason = _queuedReason;
            _queuedState = null;
            _queuedReason = string.Empty;
            SetState(queuedState, queuedReason);
        }
    }

    private Rect GetTargetRect(IslandState state)
    {
        _homePeekRect = CalculateHomePeekRect(_dockAnchor, _dockOffset);

        return state switch
        {
            IslandState.Hidden => _homePeekRect,
            IslandState.Peek => _homePeekRect,
            IslandState.Expanded => CalculateVisibleRect(_dockAnchor, ExpandedWidth, ExpandedHeightNormal),
            IslandState.NotificationOverlay => CalculateVisibleRect(_dockAnchor, ExpandedWidth, ExpandedHeightNotification),
            IslandState.Dragging => new Rect(Left, Top, Width, Height),
            _ => _homePeekRect
        };
    }

    private static Rect CalculateVisibleRect(DockAnchor anchor, double targetWidth, double targetHeight)
    {
        var area = SystemParameters.WorkArea;

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

        var hostMinHeight = state switch
        {
            IslandState.NotificationOverlay => ExpandedHeightNotification,
            IslandState.Expanded => ExpandedHeightNormal,
            _ => 0d
        };

        ContentHost.MinHeight = hostMinHeight;
        ExpandedContentHost.MinHeight = Math.Max(120, hostMinHeight);
        NotificationOverlay.MinHeight = state == IslandState.NotificationOverlay ? ExpandedHeightNotification : 0;

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
