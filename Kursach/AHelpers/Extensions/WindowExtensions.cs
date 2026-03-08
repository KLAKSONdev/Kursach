using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace Kursach.AHelpers.Extensions
{
    /// <summary>
    /// Extension методы для окон
    /// </summary>
    public static class WindowExtensions
    {


        public static void UpdateStatus(this Window window, TextBlock statusText, Ellipse indicator, string text, string colorHex)
        {
            statusText.Text = text;
            statusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));

            if (indicator != null)
                indicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
        }


        public static void SafeInvoke(this Window window, Action action)
        {
            if (window.Dispatcher.CheckAccess())
                action();
            else
                window.Dispatcher.Invoke(action);
        }
    }
}