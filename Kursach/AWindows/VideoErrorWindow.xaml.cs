using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace Kursach.AWindows
{
    public partial class VideoErrorWindow : Window
    {
        private Random rnd = new Random();
        private DispatcherTimer closeTimer;

        public VideoErrorWindow(string videoPath, int width, int height, double left, double top)
        {
            InitializeComponent();

            try
            {
                // Устанавливаем размер и позицию
                this.Width = width;
                this.Height = height;
                this.Left = left;
                this.Top = top;

                // Загружаем видео
                VideoPlayer.Source = new Uri(videoPath);
                VideoPlayer.Play();

                // Блокируем клавиши (чтобы нельзя было закрыть раньше)
                this.PreviewKeyDown += (s, e) => e.Handled = true;
                this.PreviewKeyUp += (s, e) => e.Handled = true;

                // Запрещаем закрытие пока видео не закончится
                this.Closing += (s, e) =>
                {
                    if (VideoPlayer.NaturalDuration.HasTimeSpan &&
                        VideoPlayer.Position < VideoPlayer.NaturalDuration.TimeSpan)
                    {
                        e.Cancel = true;
                    }
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка: {ex.Message}");
                this.Close();
            }
        }

        private void VideoPlayer_MediaEnded(object sender, RoutedEventArgs e)
        {
            // Даём небольшую задержку и закрываем окно
            closeTimer = new DispatcherTimer();
            closeTimer.Interval = TimeSpan.FromMilliseconds(300);
            closeTimer.Tick += (s, args) =>
            {
                closeTimer.Stop();
                this.Close();
            };
            closeTimer.Start();
        }

       
    }
}