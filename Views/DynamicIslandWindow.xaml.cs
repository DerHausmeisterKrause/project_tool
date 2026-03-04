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
    private const double NotificationHeight = 340;
    private const double Margin = 10;

    private static int _instanceCounter;

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
            ApplyVisualState(ResolveTargetState(), animate: false);
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
            {
                vm.IsExpanded = false;
                ApplyVisualState(ResolveTargetState());
            }
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
        ApplyVisualState(ResolveTargetState());
    }

    private void IslandRoot_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_isRightDragging || DataContext is not DynamicIslandViewModel vm)
            return;

        vm.ToggleExpandCommand.Execute(null);
        ApplyVisualState(ResolveTargetState());
        e.Handled = true;
    }

    private void Window_MouseEnter(object sender, MouseEventArgs e)
    {
        if (_isRightDragging)
            return;

        ApplyVisualState(ResolveTargetState());
    }

    private void Window_MouseLeave(object sender, MouseEventArgs e)
    {
        if (_isRightDragging)
            return;

        ApplyVisualState(ResolveTargetState());
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

        var dock = CalculateNearestDockPosition();
        SaveDock(dock);
        Log($"Drag end -> Snap {dock}");

        ApplyVisualState(ResolveTargetState(), animate: true);
        e.Handled = true;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(DynamicIslandViewModel.IsExpanded) or nameof(DynamicIslandViewModel.ActiveNotification))
            ApplyVisualState(ResolveTargetState());
    }

    private IslandVisualState ResolveTargetState()
    {
        if (DataContext is not DynamicIslandViewModel vm)
            return IslandVisualState.Collapsed;

        if (vm.HasNotification)
            return IslandVisualState.Notification;

        if (vm.IsExpanded)
            return IslandVisualState.Expanded;

        var hovering = IsMouseOver;
        return hovering ? IslandVisualState.Peek : IslandVisualState.Collapsed;
    }

    private DynamicIslandDockPosition CurrentDock()
    {
        var raw = ServiceLocator.Settings.Current.DynamicIslandDockPosition;
        if (!Enum.TryParse<DynamicIslandDockPosition>(raw, true, out var dock))
            dock = DynamicIslandDockPosition.TopCenter;
        return dock;
    }

    private void ApplyVisualState(IslandVisualState state, bool animate = true)
    {
        var dock = CurrentDock();
        var rect = CalculateRect(dock, state);

        if (!animate || _isRightDragging)
        {
            StopStateAnimation();
            Left = rect.Left;
            Top = rect.Top;
            Width = rect.Width;
            Height = rect.Height;
            UpdateState(state);
            return;
        }

        StopStateAnimation();
        var duration = TimeSpan.FromMilliseconds(170);

        _stateStoryboard = new Storyboard();
        AddAnimation(_stateStoryboard, this, LeftProperty, rect.Left, duration);
        AddAnimation(_stateStoryboard, this, TopProperty, rect.Top, duration);
        AddAnimation(_stateStoryboard, this, WidthProperty, rect.Width, duration);
        AddAnimation(_stateStoryboard, this, HeightProperty, rect.Height, duration);
        _stateStoryboard.Completed += (_, _) => UpdateState(state);
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

    private Rect CalculateRect(DynamicIslandDockPosition dock, IslandVisualState state)
    {
        var area = SystemParameters.WorkArea;
        var width = state switch
        {
            IslandVisualState.Collapsed => CollapsedWidth,
            IslandVisualState.Peek => PeekWidth,
            _ => ExpandedWidth
        };

        var height = state switch
        {
            IslandVisualState.Notification => NotificationHeight,
            IslandVisualState.Expanded => ExpandedHeight,
            _ => CollapsedHeight
        };

        var isExpandedState = state is IslandVisualState.Expanded or IslandVisualState.Notification;

        var topVisible = area.Top + Margin;
        var topHidden = area.Top - (height / 2);
        var bottomVisible = area.Bottom - height - Margin;
        var bottomHidden = area.Bottom - (height / 2);
        var leftVisible = area.Left + Margin;
        var leftHidden = area.Left - (width / 2);
        var rightVisible = area.Right - width - Margin;
        var rightHidden = area.Right - (width / 2);
        var centerX = area.Left + ((area.Width - width) / 2);
        var centerY = area.Top + ((area.Height - height) / 2);

        (double left, double top) = dock switch
        {
            DynamicIslandDockPosition.TopCenter => (centerX, isExpandedState ? topVisible : topHidden),
            DynamicIslandDockPosition.TopLeft => (leftVisible, isExpandedState ? topVisible : topHidden),
            DynamicIslandDockPosition.TopRight => (rightVisible, isExpandedState ? topVisible : topHidden),
            DynamicIslandDockPosition.LeftCenter => (isExpandedState ? leftVisible : leftHidden, centerY),
            DynamicIslandDockPosition.RightCenter => (isExpandedState ? rightVisible : rightHidden, centerY),
            DynamicIslandDockPosition.BottomLeft => (leftVisible, isExpandedState ? bottomVisible : bottomHidden),
            DynamicIslandDockPosition.BottomCenter => (centerX, isExpandedState ? bottomVisible : bottomHidden),
            DynamicIslandDockPosition.BottomRight => (rightVisible, isExpandedState ? bottomVisible : bottomHidden),
            _ => (centerX, isExpandedState ? topVisible : topHidden)
        };

        return new Rect(left, top, width, height);
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

    private void SaveDock(DynamicIslandDockPosition dock)
    {
        ServiceLocator.Settings.Current.DynamicIslandDockPosition = dock.ToString();
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

public enum IslandVisualState
{
    Collapsed,
    Peek,
    Expanded,
    Notification
}
