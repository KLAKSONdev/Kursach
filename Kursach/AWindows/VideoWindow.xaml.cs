using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace Kursach.AWindows
{
    public partial class VideoWindow : Window
    {
        private DispatcherTimer closeTimer;

        public VideoWindow(string videoPath)
        {
            InitializeComponent();

            try
            {
                // Загружаем видео
                VideoPlayer.Source = new Uri(videoPath);
                VideoPlayer.Play();

                // Перехватываем все клавиши
                this.PreviewKeyDown += VideoWindow_PreviewKeyDown;
                this.PreviewKeyUp += VideoWindow_PreviewKeyUp;

                // Скрываем курсор
                this.MouseMove += (s, e) => Mouse.OverrideCursor = Cursors.None;

                System.Diagnostics.Debug.WriteLine("🐕 НЕОСТАНАВЛИВАЕМЫЙ ХАСКИ АКТИВИРОВАН!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки видео: {ex.Message}");
                this.Close();
            }
        }

        // БЛОКИРУЕМ ВСЕ КЛАВИШИ
        private void VideoWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true; // Ничего не работает!
        }

        private void VideoWindow_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            e.Handled = true; // И тут тоже ничего!
        }

        // БЛОКИРУЕМ Alt+F4 и другие системные комбинации
        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            e.Handled = true;
            base.OnPreviewKeyDown(e);
        }

        // Не даём закрыть окно (закроется только когда видео закончится)
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (VideoPlayer.NaturalDuration.HasTimeSpan &&
                VideoPlayer.Position < VideoPlayer.NaturalDuration.TimeSpan)
            {
                e.Cancel = true; // НЕ ЗАКРЫВАЕМ!
            }
            base.OnClosing(e);
        }

        // Когда видео закончилось
        private void VideoPlayer_MediaEnded(object sender, RoutedEventArgs e)
        {
            // Даём время на последний кадр и закрываем
            closeTimer = new DispatcherTimer();
            closeTimer.Interval = TimeSpan.FromSeconds(0.5);
            closeTimer.Tick += (s, args) =>
            {
                closeTimer.Stop();
                this.Close();
            };
            closeTimer.Start();
        }

        // Если видео не загрузилось
        private void VideoPlayer_MediaFailed(object sender, ExceptionEventArgs e)
        {
            MessageBox.Show("Хаски сломался... 😢", "Ошибка");
            this.Close();
        }
    }
}