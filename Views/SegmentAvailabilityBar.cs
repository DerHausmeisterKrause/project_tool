using System.Collections;
using System.Windows;
using System.Windows.Media;
using TaskTool.Models;

namespace TaskTool.Views;

/// <summary>Lightweight, proportional renderer for the 06:00–18:00 availability timeline.</summary>
public sealed class SegmentAvailabilityBar : FrameworkElement
{
    private static readonly Brush FreeBrush = FrozenBrush("#22C55E");
    private static readonly Brush BusyBrush = FrozenBrush("#EF4444");
    private static readonly Brush UnknownBrush = FrozenBrush("#475569");
    private static readonly Brush LabelBrush = FrozenBrush("#94A3B8");
    private static readonly Pen SeparatorPen = FrozenPen("#0F172A", 0.7);

    public static readonly DependencyProperty SlotsProperty = DependencyProperty.Register(
        nameof(Slots), typeof(IEnumerable), typeof(SegmentAvailabilityBar),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty HasCalendarDataProperty = DependencyProperty.Register(
        nameof(HasCalendarData), typeof(bool), typeof(SegmentAvailabilityBar),
        new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));

    public IEnumerable? Slots { get => (IEnumerable?)GetValue(SlotsProperty); set => SetValue(SlotsProperty, value); }
    public bool HasCalendarData { get => (bool)GetValue(HasCalendarDataProperty); set => SetValue(HasCalendarDataProperty, value); }

    protected override Size MeasureOverride(Size availableSize) => new(double.IsInfinity(availableSize.Width) ? 0 : availableSize.Width, 34);

    protected override void OnRender(DrawingContext dc)
    {
        base.OnRender(dc);
        if (ActualWidth <= 0)
            return;

        const int slotCount = 48;
        const double barTop = 20;
        const double barHeight = 12;
        var slots = Slots?.Cast<SegmentAvailabilitySlot>().Take(slotCount).ToArray() ?? Array.Empty<SegmentAvailabilitySlot>();
        var slotWidth = ActualWidth / slotCount;

        for (var i = 0; i < slotCount; i++)
        {
            var brush = HasCalendarData && i < slots.Length
                ? (slots[i].IsBusy ? BusyBrush : FreeBrush)
                : UnknownBrush;
            var rect = new Rect(i * slotWidth, barTop, slotWidth, barHeight);
            dc.DrawRectangle(brush, i == 0 ? null : SeparatorPen, rect);
        }

        var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        for (var hour = 6; hour <= 18; hour++)
        {
            var text = new FormattedText($"{hour:00}:00", System.Globalization.CultureInfo.GetCultureInfo("de-DE"),
                FlowDirection.LeftToRight, new Typeface("Segoe UI"), 9, LabelBrush, dpi);
            var x = (hour - 6) * ActualWidth / 12d;
            var labelX = hour == 6 ? x : hour == 18 ? x - text.Width : x - text.Width / 2;
            dc.DrawText(text, new Point(Math.Max(0, labelX), 2));
            dc.DrawLine(SeparatorPen, new Point(x, 17), new Point(x, barTop));
        }
    }

    private static Brush FrozenBrush(string color)
    {
        var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));
        brush.Freeze();
        return brush;
    }

    private static Pen FrozenPen(string color, double thickness)
    {
        var pen = new Pen(FrozenBrush(color), thickness);
        pen.Freeze();
        return pen;
    }
}
