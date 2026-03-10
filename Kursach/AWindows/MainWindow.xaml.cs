using Kursach.ADialogs;
using Kursach.AHelpers.Constants;
using Kursach.AHelpers.Extensions;
using Kursach.AModels.DTO;
using Kursach.AServices;
using Kursach.AServices.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Word = Microsoft.Office.Interop.Word;

namespace Kursach.AWindows
{
    public partial class MainWindow : Window
    {
        // =====================================================================
        #region ПОЛЯ КЛАССА
        // =====================================================================

        private readonly IStudentService _studentService;
        private readonly IGroupService _groupService;
        private readonly IDocumentService _documentService;

        private List<StudentDto> _allStudents;
        private List<StudentDto> _filteredStudents;
        private List<string> _groupNames;

        public string CurrentDate => DateTime.Now.ToString("dd.MM.yyyy");
        public string UserRole { get; set; }
        public string UserName { get; set; }
        public int? UserId { get; set; }
        public int? UserGroupId { get; set; }

        #endregion

        // =====================================================================
        #region КОНСТРУКТОР
        // =====================================================================

        public MainWindow()
        {
            InitializeComponent();

            _studentService = new StudentService();
            _groupService = new GroupService();
            _documentService = new DocumentService();

            DataContext = this;
            this.PreviewKeyDown += MainWindow_PreviewKeyDown;
            Loaded += MainWindow_Loaded;
        }

        #endregion

        // =====================================================================
        #region ЗАГРУЗКА ДАННЫХ
        // =====================================================================

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                int UpdateCapsLockStatus();
                await LoadGroups();
                await LoadStudents();

                UpdateStatus($"Добро пожаловать, {UserName}!", AppColors.SuccessGreen);
            }
            catch (Exception ex)
            {
                HandleError("Ошибка загрузки данных", ex);
            }
        }

        private async System.Threading.Tasks.Task LoadGroups()
        {
            try
            {
                _groupNames = await System.Threading.Tasks.Task.Run(() =>
                    _groupService.GetGroupNames());

                await this.Dispatcher.InvokeAsync(() => UpdateGroupComboBox());
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при загрузке групп", ex);
            }
        }

        private async System.Threading.Tasks.Task LoadStudents()
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                _allStudents = await System.Threading.Tasks.Task.Run(() =>
                    _studentService.GetStudentsForRole(UserRole, UserGroupId));

                await this.Dispatcher.InvokeAsync(() =>
                {
                    _filteredStudents = new List<StudentDto>(_allStudents);
                    StudentsDataGrid.ItemsSource = _filteredStudents;

                    UpdateStatus($"Загружено {_allStudents.Count} студентов", AppColors.SuccessGreen);
                    UpdateRecordCount();
                });
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при загрузке студентов", ex);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        #endregion

        // =====================================================================
        #region ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
        // =====================================================================

        private void UpdateGroupComboBox()
        {
            GroupComboBox.Items.Clear();

            GroupComboBox.Items.Add(new ComboBoxItem
            {
                Content = "Все группы",
                IsSelected = true,
                FontWeight = FontWeights.Bold
            });

            foreach (var group in _groupNames)
            {
                GroupComboBox.Items.Add(new ComboBoxItem
                {
                    Content = group,
                    ToolTip = $"Группа {group}"
                });
            }

            StatusTextBlock.Text = $"Загружено {_groupNames.Count} групп";
        }

        private void UpdateStatus(string text, string colorHex)
        {
            StatusTextBlock.Text = text;
            StatusTextBlock.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
            StatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
        }

        private void UpdateRecordCount()
        {
            if (RecordCountText != null)
                RecordCountText.Text = $"Всего записей: {StudentsDataGrid.Items.Count}";
        }

        private void UpdateCapsLockStatus()
        {
            if (CapsLockIndicator == null) return;

            var isCapsOn = Keyboard.IsKeyToggled(Key.CapsLock);
            CapsLockIndicator.Text = isCapsOn ? "CAPS ON" : "CAPS OFF";
            CapsLockIndicator.Foreground = new SolidColorBrush(
                (Color)ColorConverter.ConvertFromString(isCapsOn ? AppColors.CapsOn : AppColors.CapsOff));
        }

        private void HandleError(string message, Exception ex)
        {
            UpdateStatus($"Ошибка: {message}", AppColors.ErrorRed);
            MessageBox.Show($"{message}: {ex.Message}", "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void ShowNotImplemented()
        {
            MessageBox.Show("Функция в разработке", "Информация",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion

        // =====================================================================
        #region ФИЛЬТРАЦИЯ
        // =====================================================================

        private void ApplyCurrentFilter()
        {
            try
            {
                var filtered = _allStudents?.AsEnumerable() ?? Enumerable.Empty<StudentDto>();

                if (GroupComboBox.SelectedItem is ComboBoxItem groupItem &&
                    groupItem.Content.ToString() != "Все группы")
                {
                    var groupName = groupItem.Content.ToString();
                    filtered = filtered.Where(s => s.GroupName == groupName);
                    CurrentGroupText.Text = $"Группа: {groupName}";
                }
                else
                {
                    CurrentGroupText.Text = "Группа: Все группы";
                }

                if (!string.IsNullOrWhiteSpace(SearchTextBox?.Text))
                {
                    var searchTerm = SearchTextBox.Text.ToLower();
                    filtered = filtered.Where(s =>
                        s.FullName?.ToLower().Contains(searchTerm) == true);
                }

                _filteredStudents = filtered.ToList();
                StudentsDataGrid.ItemsSource = _filteredStudents;

                StatusTextBlock.Text = $"Показано {_filteredStudents.Count} из {_allStudents?.Count ?? 0}";
                UpdateRecordCount();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка фильтрации: {ex.Message}");
            }
        }

        private void FilterByCourse(int course)
        {
            if (_allStudents == null) return;

            var filtered = _allStudents.Where(s => s.Course == course).ToList();
            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Показан {course} курс ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {course} курс ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        private void FilterByStatus(string status)
        {
            if (_allStudents == null) return;

            List<StudentDto> filtered;

            switch (status)
            {
                case "Выпускники":
                    filtered = _allStudents.Where(s => s.GraduationDate != null).ToList();
                    break;
                case "Архив":
                    filtered = _allStudents.Where(s => !s.IsActive).ToList();
                    break;
                default:
                    filtered = _allStudents.ToList();
                    break;
            }

            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Показаны {status} ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {status} ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        private void FilterBySocialStatus(string socialType)
        {
            if (_allStudents == null) return;

            List<StudentDto> filtered;

            switch (socialType)
            {
                case "Социально незащищенные":
                    filtered = _allStudents.Where(s => s.IsOrphan || s.IsDisabled || s.IsFromLargeFamily || s.IsLowIncome).ToList();
                    break;
                case "Сироты":
                    filtered = _allStudents.Where(s => s.IsOrphan).ToList();
                    break;
                case "Инвалиды":
                    filtered = _allStudents.Where(s => s.IsDisabled).ToList();
                    break;
                case "Многодетные семьи":
                    filtered = _allStudents.Where(s => s.IsFromLargeFamily).ToList();
                    break;
                case "Малоимущие":
                    filtered = _allStudents.Where(s => s.IsLowIncome).ToList();
                    break;
                default:
                    return;
            }

            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Показаны {socialType} ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {socialType} ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        private void FilterByAge(string ageType)
        {
            if (_allStudents == null) return;

            List<StudentDto> filtered;
            int currentYear = DateTime.Now.Year;

            switch (ageType)
            {
                case "Совершеннолетние":
                    filtered = _allStudents.Where(s => s.BirthDate.HasValue &&
                        (currentYear - s.BirthDate.Value.Year) >= 18).ToList();
                    break;
                case "Несовершеннолетние":
                    filtered = _allStudents.Where(s => s.BirthDate.HasValue &&
                        (currentYear - s.BirthDate.Value.Year) < 18).ToList();
                    break;
                default:
                    return;
            }

            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Показаны {ageType} ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {ageType} ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        #endregion

        // =====================================================================
        #region ОБРАБОТЧИКИ ИНТЕРФЕЙСА (СТАРЫЕ НАЗВАНИЯ)
        // =====================================================================

        private void GroupComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GroupComboBox.SelectedItem == null || _allStudents?.Any() != true) return;
            SearchTextBox.Text = "";
            ApplyCurrentFilter();
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyCurrentFilter();
        }

        private void StudentsDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (StudentsDataGrid.SelectedItem is StudentDto selected)
            {
                var portfolioWindow = new PortfolioWindow(selected.StudentID);
                portfolioWindow.Owner = this;
                portfolioWindow.ShowDialog();
            }
        }

        private void StudentsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StudentsDataGrid.SelectedItem is StudentDto selected)
            {
                FooterStatusText.Text = $"Выбран: {selected.FullName}";

                MilitaryCharacteristicMenuItem.Visibility =
                    (selected.Gender == "М" || selected.Gender == "Мужской")
                        ? Visibility.Visible : Visibility.Collapsed;
            }
            else
            {
                FooterStatusText.Text = "Студент не выбран";
                MilitaryCharacteristicMenuItem.Visibility = Visibility.Collapsed;
            }
        }

        // =====================================================================
        #region МЕТОДЫ ИЗ XAML (СТАРЫЕ НАЗВАНИЯ)
        // =====================================================================

        private void LoadDataButton_Click(object sender, RoutedEventArgs e)
        {
            _ = LoadGroups();
            _ = LoadStudents();
        }

        private void AddStudentButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != Roles.Administrator)
            {
                ShowNoAccess();
                return;
            }

            var dialog = new AddStudentDialog { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                _ = LoadStudents();
            }
        }

        private void EditStudent_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != Roles.Administrator)
            {
                ShowNoAccess();
                return;
            }

            if (!(StudentsDataGrid.SelectedItem is StudentDto selected))
            {
                MessageBox.Show("Выберите студента");
                return;
            }

            var dialog = new EditStudentDialog(selected.StudentID) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                _ = LoadStudents();
            }
        }

        private void DeleteStudentButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != Roles.Administrator)
            {
                ShowNoAccess();
                return;
            }

            if (!(StudentsDataGrid.SelectedItem is StudentDto selected))
            {
                MessageBox.Show("Выберите студента");
                return;
            }

            if (MessageBox.Show($"Удалить {selected.FullName}?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            try
            {
                using (var db = new vsstuEntities())
                {
                    var student = db.Students.Find(selected.StudentID);
                    if (student != null)
                    {
                        student.IsActive = false;
                        db.SaveChanges();
                        _ = LoadStudents();
                        MessageBox.Show("Студент удален", "Успех");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void Groups_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != Roles.Administrator)
            {
                ShowNoAccess();
                return;
            }

            try
            {
                this.Hide();
                var groupsWindow = new GroupsWindow();
                groupsWindow.Closed += (s, args) =>
                {
                    this.Show();
                    _ = LoadGroups();
                    _ = LoadStudents();
                };
                groupsWindow.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void StudentCards_Click(object sender, RoutedEventArgs e)
        {
            if (GroupComboBox.Items.Count > 0)
                GroupComboBox.SelectedIndex = 0;

            StudentsDataGrid.ItemsSource = _allStudents;
            StatusTextBlock.Text = $"Показаны все студенты ({_allStudents?.Count ?? 0} чел.)";
            FooterStatusText.Text = $"Все студенты ({_allStudents?.Count ?? 0} записей)";
            CurrentGroupText.Text = "Группа: Все группы";
            UpdateRecordCount();
        }

        private void Registration_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is MenuItem menuItem)) return;

            string header = menuItem.Header.ToString();

            if (header.Contains("1 курс")) FilterByCourse(1);
            else if (header.Contains("2 курс")) FilterByCourse(2);
            else if (header.Contains("3 курс")) FilterByCourse(3);
            else if (header.Contains("4 курс")) FilterByCourse(4);
            else if (header == "Выпускники") FilterByStatus("Выпускники");
            else if (header == "Архив") FilterByStatus("Архив");
            else if (header == "Социально незащищенные") FilterBySocialStatus("Социально незащищенные");
            else if (header == "Сироты") FilterBySocialStatus("Сироты");
            else if (header == "Инвалиды") FilterBySocialStatus("Инвалиды");
            else if (header == "Многодетные семьи") FilterBySocialStatus("Многодетные семьи");
            else if (header == "Малоимущие") FilterBySocialStatus("Малоимущие");
            else if (header == "Совершеннолетние") FilterByAge("Совершеннолетние");
            else if (header == "Несовершеннолетние") FilterByAge("Несовершеннолетние");
        }

        private void EventsAdmin_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var eventsWindow = new EventsWindow(isAdminMode: true) { Owner = this };
                eventsWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void EventsStudent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var eventsWindow = new EventsWindow(isAdminMode: false, groupId: UserGroupId) { Owner = this };
                eventsWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void CreateStudyCertificate_Click(object sender, RoutedEventArgs e)
        {
            if (!(StudentsDataGrid.SelectedItem is StudentDto selected))
            {
                MessageBox.Show("Выберите студента");
                return;
            }

            _documentService.CreateStudyCertificate(selected);
        }

        private void CreateMilitaryCharacteristic_Click(object sender, RoutedEventArgs e)
        {
            if (!(StudentsDataGrid.SelectedItem is StudentDto selected))
            {
                MessageBox.Show("Выберите студента");
                return;
            }

            _documentService.CreateMilitaryCharacteristic(selected);
        }

        private void GenerateCharacteristic_Click(object sender, RoutedEventArgs e)
        {
            ShowNotImplemented();
        }

        private void GenerateSocialPassport_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        private void GeneratePortfolio_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        private void ExportToExcel_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        private void ExportToWord_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();

        private void ShowNoAccess()
        {
            MessageBox.Show("У вас нет прав для этого действия", "Доступ запрещен",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        #endregion

        // =====================================================================
        #region ПАСХАЛКИ
        // =====================================================================

        private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.H && Keyboard.IsKeyDown(Key.LeftCtrl))
            {
                e.Handled = true;
                OpenHaskiVideo();
            }
        }

        private void OpenHaskiVideo()
        {
            try
            {
                string videoPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "AVideos",
                    "DJhvost.mp4"
                );

                if (!File.Exists(videoPath))
                {
                    videoPath = Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        @"..\..\AVideos\DJhvost.mp4"
                    );
                    videoPath = Path.GetFullPath(videoPath);
                }

                if (!File.Exists(videoPath))
                {
                    MessageBox.Show("Видео с Хаски не найдено, бро! 🙁\nПоложи файл haski.mp4 в папку Videos",
                                  "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var videoWindow = new VideoWindow(videoPath) { Owner = this };
                videoWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии видео: {ex.Message}");
            }
        }

        private void DimaButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Лососни тунца", "Лапух",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            OpenDimaTrojan();
        }

        private void OpenDimaTrojan()
        {
            try
            {
                string videoPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "AVideos",
                    "DJhvost.mp4"
                );

                if (!File.Exists(videoPath))
                {
                    MessageBox.Show("Видео с Хаски не найдено!");
                    return;
                }

                Random rnd = new Random();
                int screenWidth = (int)SystemParameters.PrimaryScreenWidth;
                int screenHeight = (int)SystemParameters.PrimaryScreenHeight;

                for (int i = 0; i < 15; i++)
                {
                    int width = rnd.Next(300, 600);
                    int height = rnd.Next(250, 450);
                    int left = rnd.Next(0, screenWidth - width);
                    int top = rnd.Next(0, screenHeight - height);

                    var videoWindow = new VideoErrorWindow(videoPath, width, height, left, top);
                    videoWindow.Show();

                    System.Threading.Thread.Sleep(150);
                }

                this.WindowState = WindowState.Minimized;
                this.ShowInTaskbar = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        #endregion
        #endregion
    }
}
