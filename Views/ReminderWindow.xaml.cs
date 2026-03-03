using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using TaskTool.Services;

namespace TaskTool.Views;

public partial class ReminderWindow : Window
{
    private readonly DispatcherTimer _hideTimer;
    private readonly Guid _taskId;

    public event EventHandler<Guid>? NotificationClicked;

    public ReminderWindow(string text, ReminderKind kind, Guid taskId)
    {
        InitializeComponent();

        _taskId = taskId;
        MessageText.Text = text;
        PillBorder.BorderBrush = kind == ReminderKind.Start
            ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239, 68, 68))
            : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(245, 158, 11));

        Loaded += OnLoaded;
        MouseLeftButtonUp += OnMouseLeftButtonUp;

        _hideTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(8) };
        _hideTimer.Tick += (_, _) => BeginHide();
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        PositionTopCenter();
        BeginShow();
        _hideTimer.Start();
    }

    private void PositionTopCenter()
    {
        Left = SystemParameters.WorkArea.Left + (SystemParameters.WorkArea.Width - Width) / 2;
        Top = SystemParameters.WorkArea.Top + 12;
    }

    private void BeginShow()
    {
        var sb = new Storyboard();

        var fade = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(220));
        Storyboard.SetTarget(fade, PillBorder);
        Storyboard.SetTargetProperty(fade, new PropertyPath(Border.OpacityProperty));

        var slide = new DoubleAnimation(Top - 8, Top, TimeSpan.FromMilliseconds(220));
        Storyboard.SetTarget(slide, this);
        Storyboard.SetTargetProperty(slide, new PropertyPath(Window.TopProperty));

        sb.Children.Add(fade);
        sb.Children.Add(slide);
        sb.Begin();
    }

    private void BeginHide()
    {
        _hideTimer.Stop();
        var sb = new Storyboard();

        var fade = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(220));
        Storyboard.SetTarget(fade, PillBorder);
        Storyboard.SetTargetProperty(fade, new PropertyPath(Border.OpacityProperty));

        var slide = new DoubleAnimation(Top, Top - 8, TimeSpan.FromMilliseconds(220));
        Storyboard.SetTarget(slide, this);
        Storyboard.SetTargetProperty(slide, new PropertyPath(Window.TopProperty));

        sb.Children.Add(fade);
        sb.Children.Add(slide);
        sb.Completed += (_, _) => Close();
        sb.Begin();
    }

    private void OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        NotificationClicked?.Invoke(this, _taskId);
        BeginHide();
    }
}
