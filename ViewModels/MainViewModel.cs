using System;
using System.Windows;
using System.Collections.ObjectModel;
using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class MainViewModel : ObservableObject
{
    public ObservableCollection<object> NavigationItems { get; }
    public TodayViewModel TodayViewModel { get; }
    private readonly WeekViewModel _weekViewModel;
    private readonly TicketSystemViewModel _ticketSystemViewModel;
    public WikiBrowserViewModel WikiBrowserViewModel { get; }
    private readonly ReportsViewModel _reportsViewModel;
    private readonly List<WebShortcutViewModel> _webShortcutViews = new();
    private readonly LoggerService _logger;
    public SettingsViewModel SettingsViewModel { get; }
    public RelayCommand NavigateToSettingsCommand { get; }

    private object _selectedView;
    public object SelectedView
    {
        get => _selectedView;
        set
        {
            if (Set(ref _selectedView, value))
            {
                if (_selectedView is WeekViewModel)
                {
                    _weekViewModel.Refresh();
                }
                else if (_selectedView is TicketSystemViewModel)
                {
                    _ticketSystemViewModel.Refresh();
                }
                else if (_selectedView is ReportsViewModel)
                {
                    _reportsViewModel.Refresh();
                }
                else if (_selectedView is WikiBrowserViewModel) WikiBrowserViewModel.EnsureHome();
                Raise(nameof(IsTodaySelected));
            }
        }
    }

    public bool IsTodaySelected => SelectedView is TodayViewModel;

    public event Action? FocusQuickAddRequested;

    public void NavigateToTodayAndOpenTask(Guid taskId)
    {
        SelectedView = TodayViewModel;

        if (TodayViewModel.NavigateToTask(taskId))
            return;

        MessageBox.Show("Aufgabe nicht gefunden", "Kalender", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    public void NavigateToTodayAndFocusQuickAdd()
    {
        SelectedView = TodayViewModel;
        FocusQuickAddRequested?.Invoke();
    }

    public void NavigateToTodayCurrentTasks()
    {
        SelectedView = TodayViewModel;
        TodayViewModel.SelectedTaskScope = TodayTaskScope.Current;
    }

    public void NavigateToTicketSystem(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return;

        _ticketSystemViewModel.NavigateTo(uri.ToString());
        SelectedView = _ticketSystemViewModel;
        var ticketId = System.Text.RegularExpressions.Regex.Match(uri.Query, @"(?:[?;&])TicketID=([^;&]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups[1].Value;
        _logger.Info($"[TicketOpenInApp] ticketId='{Uri.UnescapeDataString(ticketId)}' ticketNumber='' targetUrl='{uri.Scheme}://{uri.Host}{uri.AbsolutePath}' targetTab=Ticketsystem");
    }

    public void NavigateToWiki(string sourceId, string url)
    {
        RefreshWikiNavigation();
        WikiBrowserViewModel.NavigateTo(sourceId, url);
        if (NavigationItems.Contains(WikiBrowserViewModel)) SelectedView = WikiBrowserViewModel;
    }

    public MainViewModel(TaskService taskService, WorkDayService workDayService, SettingsService settingsService, NotificationService notifications, OutlookCalendarService outlookCalendar, TicketSystemService ticketSystem, UpdateService updates, HomeOfficeService homeOffice, GermanTimeService germanTime, LoggerService logger)
    {
        _logger = logger;
        TodayViewModel = new TodayViewModel(taskService, workDayService, settingsService, outlookCalendar, ticketSystem, homeOffice);
        _weekViewModel = new WeekViewModel(taskService, workDayService, settingsService, outlookCalendar, homeOffice);
        ticketSystem.TasksChanged += TodayViewModel.Refresh;
        ticketSystem.TasksChanged += _weekViewModel.Refresh;
        _ticketSystemViewModel = new TicketSystemViewModel(settingsService);
        WikiBrowserViewModel = new WikiBrowserViewModel(settingsService);
        _reportsViewModel = new ReportsViewModel(taskService, workDayService, settingsService, germanTime, logger);
        SettingsViewModel = new SettingsViewModel(settingsService, notifications, outlookCalendar, taskService, ticketSystem, updates);
        NavigateToSettingsCommand = new RelayCommand(() => SelectedView = SettingsViewModel);

        NavigationItems = new ObservableCollection<object> { TodayViewModel, _weekViewModel, _ticketSystemViewModel, _reportsViewModel, SettingsViewModel };
        settingsService.SettingsChanged += RefreshDynamicNavigation;
        RefreshDynamicNavigation();
        _selectedView = TodayViewModel;
    }

    private void RefreshWikiNavigation()
    {
        WikiBrowserViewModel.RefreshSources();
        var shouldShow = WikiBrowserViewModel.Sources.Count > 0;
        var contains = NavigationItems.Contains(WikiBrowserViewModel);
        if (shouldShow && !contains) NavigationItems.Insert(NavigationItems.IndexOf(_ticketSystemViewModel) + 1, WikiBrowserViewModel);
        else if (!shouldShow && contains)
        {
            NavigationItems.Remove(WikiBrowserViewModel);
            if (ReferenceEquals(SelectedView, WikiBrowserViewModel)) SelectedView = TodayViewModel;
        }
    }
    private void RefreshDynamicNavigation()
    {
        RefreshWikiNavigation();
        var selectedShortcut = SelectedView as WebShortcutViewModel;
        foreach (var view in _webShortcutViews) NavigationItems.Remove(view);
        _webShortcutViews.Clear();
        foreach (var shortcut in ServiceLocator.Settings.Current.WebShortcuts.Where(x => x.Enabled))
        {
            var view = new WebShortcutViewModel(shortcut);
            _webShortcutViews.Add(view);
            NavigationItems.Insert(NavigationItems.IndexOf(SettingsViewModel), view);
        }
        if (selectedShortcut != null)
        {
            var replacement = _webShortcutViews.FirstOrDefault(x => x.ShortcutId == selectedShortcut.ShortcutId);
            SelectedView = replacement is not null
                ? replacement
                : TodayViewModel;
        }
    }
}
