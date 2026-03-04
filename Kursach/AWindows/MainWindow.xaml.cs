using Kursach.ADialogs;
using Kursach.AModels;
using Kursach.AWindows;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
        public string CurrentDate => DateTime.Now.ToString("dd.MM.yyyy");

        private List<StudentViewModel> allStudents = new List<StudentViewModel>();
        private List<string> groupNames = new List<string>();

        public string UserRole { get; set; }
        public string UserName { get; set; }
        public int? UserId { get; set; }
        public int? UserGroupId { get; set; }

        private EventsWindow eventsWindow = null;
        #endregion

        // =====================================================================
        #region КОНСТРУКТОР
        // =====================================================================
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            Loaded += MainWindow_Loaded;
        }
        #endregion

        // =====================================================================
        #region ЗАГРУЗКА ДАННЫХ
        // =====================================================================
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateCapsLockStatus();
                LoadGroups();
                LoadStudents();

                UpdateStatus($"{UserName}", "#27AE60");
            }
            catch (Exception ex)
            {
                UpdateStatus("Статус: Ошибка загрузки", "#E74C3C");
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadGroups()
        {
            try
            {
                using (var db = new vsstuEntities())
                {
                    groupNames = db.Groups
                        .Where(g => g.Students.Any(s => s.IsActive == true))
                        .Select(g => g.GroupName)
                        .Distinct()
                        .OrderBy(g => g)
                        .ToList();

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        GroupComboBox.Items.Clear();

                        var allGroupsItem = new ComboBoxItem
                        {
                            Content = "Все группы",
                            IsSelected = true,
                            FontWeight = FontWeights.Bold
                        };
                        GroupComboBox.Items.Add(allGroupsItem);

                        foreach (var group in groupNames)
                        {
                            GroupComboBox.Items.Add(new ComboBoxItem
                            {
                                Content = group,
                                ToolTip = $"Группа {group}"
                            });
                        }

                        if (groupNames.Any())
                        {
                            StatusTextBlock.Text = $"Статус: Загружено {groupNames.Count} групп";
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadStudents()
        {
            try
            {
                using (var db = new vsstuEntities())
                {
                    IQueryable<Students> query = db.Students.Where(s => s.IsActive == true);

                    if (UserRole == "Староста" && UserGroupId.HasValue)
                    {
                        query = query.Where(s => s.GroupID == UserGroupId.Value);
                    }

                    var dbStudents = query.ToList();

                    allStudents = new List<StudentViewModel>();

                    foreach (var s in dbStudents)
                    {
                        var student = new StudentViewModel
                        {
                            StudentID = s.StudentID,
                            LastName = s.LastName ?? "",
                            FirstName = s.FirstName ?? "",
                            MiddleName = s.MiddleName ?? "",
                            GroupID = s.GroupID,
                            GroupName = s.Groups?.GroupName ?? "Без группы",
                            Course = s.Groups?.Course,
                            StudentCardNumber = s.StudentCardNumber ?? "",
                            PersonalNumber = s.PersonalNumber ?? "",
                            BirthDate = s.BirthDate,
                            BirthPlace = s.BirthPlace ?? "",
                            Gender = s.Gender ?? "",
                            Nationality = s.Nationality ?? "",
                            Citizenship = s.Citizenship ?? "",
                            Phone = s.Phone ?? "",
                            Email = s.Email ?? "",
                            ParentsPhone = s.ParentsPhone ?? "",
                            RegistrationAddress = s.RegistrationAddress ?? "",
                            ResidentialAddress = s.ResidentialAddress ?? "",
                            IsOrphan = s.IsOrphan ?? false,
                            IsDisabled = s.IsDisabled ?? false,
                            IsFromLargeFamily = s.IsFromLargeFamily ?? false,
                            IsLowIncome = s.IsLowIncome ?? false,
                            IsEmployed = s.IsEmployed ?? false,
                            WorkPlace = s.WorkPlace ?? "",
                            WorkPosition = s.WorkPosition ?? "",
                            Login = s.Login ?? "",
                            IsActive = s.IsActive ?? true,
                            EnrollmentDate = s.EnrollmentDate,
                            GraduationDate = s.GraduationDate,
                            CreatedAt = s.CreatedAt,
                            UpdatedAt = s.UpdatedAt,
                            IsHeadman = s.IsHeadman ?? false
                        };

                        student.FullName = $"{student.LastName} {student.FirstName} {student.MiddleName}".Trim();

                        if (s.BirthDate.HasValue)
                        {
                            student.Age = DateTime.Now.Year - s.BirthDate.Value.Year;
                            if (DateTime.Now.DayOfYear < s.BirthDate.Value.DayOfYear)
                                student.Age--;
                        }

                        var statuses = new List<string>();
                        if (s.IsOrphan == true) statuses.Add("Сирота");
                        if (s.IsDisabled == true) statuses.Add("Инвалид");
                        if (s.IsFromLargeFamily == true) statuses.Add("Многодетная семья");
                        if (s.IsLowIncome == true) statuses.Add("Малообеспеченный");
                        student.SocialStatus = statuses.Count > 0 ? string.Join(", ", statuses) : "Не указан";

                        allStudents.Add(student);
                    }

                    StudentsDataGrid.ItemsSource = allStudents;
                    UpdateStatus($"Статус: Загружено {allStudents.Count} студентов", "#27AE60");
                    UpdateRecordCount();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки: {ex.Message}", "Ошибка");
            }
        }
        #endregion

        // =====================================================================
        #region ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
        // =====================================================================
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
            if (CapsLockIndicator != null)
            {
                CapsLockIndicator.Text = Keyboard.IsKeyToggled(Key.CapsLock) ? "CAPS ON" : "CAPS OFF";
                CapsLockIndicator.Foreground = Keyboard.IsKeyToggled(Key.CapsLock) ?
                    new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E74C3C")) :
                    new SolidColorBrush((Color)ColorConverter.ConvertFromString("#808080"));
            }
        }

        private void ShowNotImplemented()
        {
            MessageBox.Show("Функция в разработке", "Информация",
                          MessageBoxButton.OK, MessageBoxImage.Information);
        }
        #endregion

        // =====================================================================
        #region ФИЛЬТРАЦИЯ ДАННЫХ
        // =====================================================================
        private void FilterByCourse(int course)
        {
            var filtered = allStudents.Where(s => s.Course == course).ToList();
            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Статус: Показан {course} курс ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {course} курс ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        private void FilterByStatus(string status)
        {
            List<StudentViewModel> filtered = new List<StudentViewModel>();

            switch (status)
            {
                case "Выпускники":
                    filtered = allStudents.Where(s => s.GraduationDate != null).ToList();
                    break;
                case "Архив":
                    filtered = allStudents.Where(s => s.IsActive == false).ToList();
                    break;
                default:
                    filtered = allStudents;
                    break;
            }

            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Статус: Показаны {status} ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {status} ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        private void FilterBySocialStatus(string socialType)
        {
            List<StudentViewModel> filtered = new List<StudentViewModel>();

            switch (socialType)
            {
                case "Социально незащищенные":
                    filtered = allStudents.Where(s => s.IsOrphan || s.IsDisabled || s.IsFromLargeFamily || s.IsLowIncome).ToList();
                    break;
                case "Сироты":
                    filtered = allStudents.Where(s => s.IsOrphan).ToList();
                    break;
                case "Инвалиды":
                    filtered = allStudents.Where(s => s.IsDisabled).ToList();
                    break;
                case "Многодетные семьи":
                    filtered = allStudents.Where(s => s.IsFromLargeFamily).ToList();
                    break;
                case "Малоимущие":
                    filtered = allStudents.Where(s => s.IsLowIncome).ToList();
                    break;
                default:
                    return;
            }

            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Статус: Показаны {socialType} ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {socialType} ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        private void FilterByAge(string ageType)
        {
            List<StudentViewModel> filtered = new List<StudentViewModel>();
            int currentYear = DateTime.Now.Year;

            switch (ageType)
            {
                case "Совершеннолетние":
                    filtered = allStudents.Where(s => s.BirthDate.HasValue &&
                        (currentYear - s.BirthDate.Value.Year) >= 18).ToList();
                    break;
                case "Несовершеннолетние":
                    filtered = allStudents.Where(s => s.BirthDate.HasValue &&
                        (currentYear - s.BirthDate.Value.Year) < 18).ToList();
                    break;
                default:
                    return;
            }

            StudentsDataGrid.ItemsSource = filtered;
            StatusTextBlock.Text = $"Статус: Показаны {ageType} ({filtered.Count} чел.)";
            FooterStatusText.Text = $"Фильтр: {ageType} ({filtered.Count} записей)";
            UpdateRecordCount();
        }

        private void ApplyCurrentFilter()
        {
            try
            {
                var selectedItem = GroupComboBox.SelectedItem as ComboBoxItem;
                if (selectedItem == null) return;

                string selectedGroup = selectedItem.Content.ToString();
                List<StudentViewModel> filtered;

                if (selectedGroup == "Все группы")
                {
                    filtered = allStudents.ToList();
                    CurrentGroupText.Text = "Группа: Все группы";
                }
                else
                {
                    filtered = allStudents.Where(s => s.GroupName == selectedGroup).ToList();
                    if (filtered.Count > 0 && filtered.First().Course.HasValue)
                    {
                        CurrentGroupText.Text = $"Группа: {selectedGroup} ({filtered.First().Course} курс)";
                    }
                    else
                    {
                        CurrentGroupText.Text = $"Группа: {selectedGroup}";
                    }
                }

                if (!string.IsNullOrWhiteSpace(SearchTextBox?.Text))
                {
                    string searchText = SearchTextBox.Text.ToLower().Trim();
                    filtered = filtered.Where(s =>
                        (s.LastName != null && s.LastName.ToLower().Contains(searchText)) ||
                        (s.FirstName != null && s.FirstName.ToLower().Contains(searchText)) ||
                        (s.MiddleName != null && s.MiddleName.ToLower().Contains(searchText)) ||
                        (s.FullName != null && s.FullName.ToLower().Contains(searchText))
                    ).ToList();
                }

                StudentsDataGrid.ItemsSource = filtered;
                StatusTextBlock.Text = $"Статус: Показано {filtered.Count} студентов";
                UpdateRecordCount();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка фильтрации: {ex.Message}");
            }
        }
        #endregion

        // =====================================================================
        #region ОБРАБОТЧИКИ ИНТЕРФЕЙСА
        // =====================================================================
        private void GroupComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (GroupComboBox.SelectedItem == null || allStudents.Count == 0) return;
                SearchTextBox.Text = "";
                ApplyCurrentFilter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при фильтрации: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void StudentsDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (StudentsDataGrid.SelectedItem == null) return;

                var selectedStudent = StudentsDataGrid.SelectedItem as StudentViewModel;
                if (selectedStudent == null) return;

                var portfolioWindow = new PortfolioWindow(selectedStudent.StudentID);
                portfolioWindow.Owner = this;
                portfolioWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии портфолио: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void StudentsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StudentsDataGrid.SelectedItem != null)
            {
                var selected = StudentsDataGrid.SelectedItem as StudentViewModel;
                if (selected != null)
                {
                    FooterStatusText.Text = $"Выбран: {selected.LastName} {selected.FirstName}";

                    MilitaryCharacteristicMenuItem.Visibility = (selected.Gender == "М" || selected.Gender == "Мужской")
                        ? Visibility.Visible : Visibility.Collapsed;
                }
            }
            else
            {
                FooterStatusText.Text = "Студент не выбран";
                MilitaryCharacteristicMenuItem.Visibility = Visibility.Collapsed;
            }
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyCurrentFilter();
        }
        #endregion

        // =====================================================================
        #region ОБРАБОТЧИКИ КНОПОК
        // =====================================================================
        private void EditStudent_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != "Администратор")
            {
                MessageBox.Show("У вас нет прав для редактирования", "Доступ запрещен",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (StudentsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите студента для редактирования", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selected = StudentsDataGrid.SelectedItem as StudentViewModel;
            if (selected == null) return;

            try
            {
                var dialog = new ADialogs.EditStudentDialog(selected.StudentID);
                dialog.Owner = this;

                if (dialog.ShowDialog() == true)
                {
                    LoadGroups();
                    LoadStudents();
                    MessageBox.Show("Данные студента обновлены", "Успех",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии редактора: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDataButton_Click(object sender, RoutedEventArgs e)
        {
            LoadGroups();
            LoadStudents();
        }

        private void AddStudentButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != "Администратор")
            {
                MessageBox.Show("У вас нет прав для добавления студентов", "Доступ запрещен",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var dialog = new ADialogs.AddStudentDialog();
            dialog.Owner = this;
            if (dialog.ShowDialog() == true)
            {
                LoadGroups();
                LoadStudents();
            }
        }

        private void DeleteStudentButton_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != "Администратор")
            {
                MessageBox.Show("У вас нет прав для удаления", "Доступ запрещен",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (StudentsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите студента для удаления", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selected = StudentsDataGrid.SelectedItem as StudentViewModel;
            if (selected == null) return;

            string fullName = $"{selected.LastName} {selected.FirstName} {selected.MiddleName}".Trim();

            if (MessageBox.Show($"Удалить {fullName}?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    using (var db = new vsstuEntities())
                    {
                        var student = db.Students.Find(selected.StudentID);
                        if (student != null)
                        {
                            student.IsActive = false;
                            db.SaveChanges();
                            MessageBox.Show("Студент удален", "Успех",
                                          MessageBoxButton.OK, MessageBoxImage.Information);
                            LoadGroups();
                            LoadStudents();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void TestConnectionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var db = new vsstuEntities())
                {
                    if (db.Database.Exists())
                    {
                        UpdateStatus("Статус: Подключение к БД успешно", "#27AE60");
                        FooterStatusText.Text = "Подключено к базе данных";
                    }
                    else
                    {
                        UpdateStatus("Статус: База данных не найдена", "#E74C3C");
                        FooterStatusText.Text = "БД не найдена";
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Статус: Ошибка подключения", "#E74C3C");
                FooterStatusText.Text = "Ошибка подключения";
                MessageBox.Show($"Ошибка подключения к БД: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        // =====================================================================
        #region ДОКУМЕНТЫ
        // =====================================================================
        private void CreateStudyCertificate_Click(object sender, RoutedEventArgs e)
        {
            if (StudentsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите студента", "Информация");
                return;
            }

            var selected = StudentsDataGrid.SelectedItem as StudentViewModel;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word документы (*.docx)|*.docx";
            saveFileDialog.FileName = $"Справка_об_обучении_{selected.LastName}_{DateTime.Now:yyyyMMdd}.docx";

            if (saveFileDialog.ShowDialog() == true)
            {
                CreateStudyCertificateDocument(saveFileDialog.FileName, selected);
            }
        }

        private void CreateStudyCertificateDocument(string filePath, StudentViewModel student)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                doc = wordApp.Documents.Add();
                Word.Selection selection = wordApp.Selection;

                selection.Font.Name = "Times New Roman";
                selection.Font.Size = 14;

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Bold = 1;
                selection.Font.Size = 16;
                selection.TypeText("СПРАВКА ОБ ОБУЧЕНИИ\n");
                selection.Font.Bold = 0;
                selection.Font.Size = 14;
                selection.TypeText("\n\n");

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                selection.TypeText($"Дана {student.LastName} {student.FirstName} {student.MiddleName},\n");
                selection.TypeText($"студенту группы {student.GroupName},\n");
                selection.TypeText($"в том, что он(а) действительно обучается в\n");
                selection.TypeText("ГБПОУ \"Колледж\" по специальности\n");
                selection.TypeText($"{GetStudentSpecialty(student.StudentID)}.\n\n");
                selection.TypeText($"Курс: {student.Course}\n");
                selection.TypeText($"Форма обучения: очная\n\n");
                selection.TypeText("Справка выдана для предъявления по месту требования.\n\n\n\n");

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                selection.TypeText("Директор колледжа _____________ И.И. Иванов\n");
                selection.TypeText("М.П.\n\n");
                selection.TypeText($"«___» __________ {DateTime.Now.Year} г.");

                doc.SaveAs2(filePath);
                MessageBox.Show("Справка об обучении создана!", "Успех");
            }
            finally
            {
                if (doc != null) { doc.Close(); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
            }
        }

        private void CreateAcademicLeaveApplication_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (StudentsDataGrid.SelectedItem == null)
                {
                    MessageBox.Show("Выберите студента из списка", "Информация");
                    return;
                }

                var selectedStudent = StudentsDataGrid.SelectedItem as StudentViewModel;
                if (selectedStudent == null) return;

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word документы (*.docx)|*.docx";
                saveFileDialog.DefaultExt = "docx";
                saveFileDialog.FileName = $"Заявление_академ_{selectedStudent.LastName}_{DateTime.Now:yyyyMMdd}.docx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    CreateAcademicLeaveDocument(saveFileDialog.FileName, selectedStudent);
                    MessageBox.Show("Заявление успешно создано!", "Успех");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private void CreateAcademicLeaveDocument(string filePath, StudentViewModel student)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                doc = wordApp.Documents.Add();
                Word.Selection selection = wordApp.Selection;

                selection.Font.Name = "Times New Roman";
                selection.Font.Size = 14;

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                selection.TypeText("Директору колледжа\n");
                selection.TypeText("Иванову И.И.\n");
                selection.TypeText($"от студента группы {student.GroupName}\n");
                selection.TypeText($"{student.LastName} {student.FirstName} {student.MiddleName}\n");
                selection.TypeText("\n");

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Bold = 1;
                selection.Font.Size = 16;
                selection.TypeText("ЗАЯВЛЕНИЕ\n");
                selection.Font.Bold = 0;
                selection.Font.Size = 14;
                selection.TypeText("\n");

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                selection.TypeText("Прошу предоставить мне академический отпуск по семейным обстоятельствам.\n\n");
                selection.TypeText("Срок отпуска: с «___» __________ 20__ г. по «___» __________ 20__ г.\n\n");
                selection.TypeText("Обязуюсь предоставить подтверждающие документы.\n\n");
                selection.TypeText("«___» __________ 20___ г.\n\n");
                selection.TypeText("__________________\n");
                selection.TypeText("     (подпись)\n");

                doc.SaveAs2(filePath);
            }
            finally
            {
                if (doc != null) { doc.Close(); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
            }
        }

        private void CreateMilitaryCharacteristic_Click(object sender, RoutedEventArgs e)
        {
            if (StudentsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите студента", "Информация");
                return;
            }

            var selected = StudentsDataGrid.SelectedItem as StudentViewModel;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word документы (*.docx)|*.docx";
            saveFileDialog.FileName = $"Характеристика_в_военкомат_{selected.LastName}_{DateTime.Now:yyyyMMdd}.docx";

            if (saveFileDialog.ShowDialog() == true)
            {
                CreateMilitaryCharacteristicDocument(saveFileDialog.FileName, selected);
            }
        }

        private void CreateMilitaryCharacteristicDocument(string filePath, StudentViewModel student)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                doc = wordApp.Documents.Add();
                Word.Selection selection = wordApp.Selection;

                selection.Font.Name = "Times New Roman";
                selection.Font.Size = 14;

                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Bold = 1;
                selection.Font.Size = 16;
                selection.TypeText("ХАРАКТЕРИСТИКА\n\n");
                selection.Font.Bold = 0;
                selection.Font.Size = 14;

                selection.TypeText($"на студента {student.LastName} {student.FirstName} {student.MiddleName}\n");
                selection.TypeText($"{student.BirthDate:dd.MM.yyyy} года рождения\n\n");

                selection.Font.Bold = 1;
                selection.TypeText("1. УСПЕВАЕМОСТЬ\n");
                selection.Font.Bold = 0;
                selection.TypeText($"За время обучения проявил себя как {GetAcademicPerformanceText(student.StudentID)} студент. ");
                selection.TypeText($"Средний балл успеваемости: {GetAverageGrade(student.StudentID)}.\n\n");

                selection.Font.Bold = 1;
                selection.TypeText("2. ДИСЦИПЛИНА\n");
                selection.Font.Bold = 0;
                selection.TypeText($"Правила внутреннего распорядка {GetDisciplineText(student.StudentID)}. ");
                selection.TypeText($"Замечаний от преподавателей {GetRemarksText(student.StudentID)}.\n\n");

                selection.Font.Bold = 1;
                selection.TypeText("3. ОБЩЕСТВЕННАЯ АКТИВНОСТЬ\n");
                selection.Font.Bold = 0;
                selection.TypeText($"В общественной жизни группы и колледжа {GetActivityText(student.StudentID)}. ");
                selection.TypeText($"Количество мероприятий с участием: {GetEventsCount(student.StudentID)}.\n\n");

                selection.Font.Bold = 1;
                selection.TypeText("4. ЛИЧНЫЕ КАЧЕСТВА\n");
                selection.Font.Bold = 0;
                selection.TypeText(GetPersonalQualities(student.StudentID) + "\n\n");

                selection.TypeText("Куратор группы __________________ /Иванова М.И./\n\n");
                selection.TypeText("Директор колледжа __________________ /Иванов И.И./\n");
                selection.TypeText("М.П.\n\n");
                selection.TypeText($"«___» __________ {DateTime.Now.Year} г.\n");

                doc.SaveAs2(filePath);
                MessageBox.Show("Характеристика в военкомат создана!", "Успех");
            }
            finally
            {
                if (doc != null) { doc.Close(); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
            }
        }
        #endregion

        // =====================================================================
        #region ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ДЛЯ ДОКУМЕНТОВ
        // =====================================================================
        private string GetStudentSpecialty(int studentId)
        {
            using (var db = new vsstuEntities())
            {
                var student = db.Students
                    .FirstOrDefault(s => s.StudentID == studentId);
                return student?.Groups?.Specialties?.SpecialtyName ?? "Информационные технологии";
            }
        }

        private double GetAverageGrade(int studentId)
        {
            using (var db = new vsstuEntities())
            {
                var grades = db.AcademicPerformance
                    .Where(g => g.StudentID == studentId && g.Grade.HasValue)
                    .Select(g => g.Grade.Value);
                return grades.Any() ? Math.Round(grades.Average(), 2) : 4.0;
            }
        }

        private string GetAcademicPerformanceText(int studentId)
        {
            double avg = GetAverageGrade(studentId);
            if (avg >= 4.5) return "отличный";
            if (avg >= 3.5) return "хороший";
            if (avg >= 2.5) return "удовлетворительный";
            return "слабый";
        }

        private string GetDisciplineText(int studentId)
        {
            using (var db = new vsstuEntities())
            {
                int violations = db.DisciplinaryRecords
                    .Count(d => d.StudentID == studentId);
                return violations == 0 ? "соблюдает" : "иногда нарушает";
            }
        }

        private string GetRemarksText(int studentId)
        {
            using (var db = new vsstuEntities())
            {
                int remarks = db.DisciplinaryRecords
                    .Count(d => d.StudentID == studentId && d.RecordType == "Замечание");
                return remarks == 0 ? "не имеет" : $"имеет ({remarks})";
            }
        }

        private string GetActivityText(int studentId)
        {
            using (var db = new vsstuEntities())
            {
                int events = db.EventParticipation
                    .Count(ep => ep.StudentID == studentId);
                if (events > 10) return "принимает активное участие";
                if (events > 5) return "участвует";
                if (events > 0) return "эпизодически участвует";
                return "не участвует";
            }
        }

        private int GetEventsCount(int studentId)
        {
            using (var db = new vsstuEntities())
            {
                return db.EventParticipation.Count(ep => ep.StudentID == studentId);
            }
        }

        private string GetPersonalQualities(int studentId)
        {
            using (var db = new vsstuEntities())
            {
                var traits = db.StudentTraits
                    .Where(st => st.StudentID == studentId)
                    .ToList();

                var positive = traits
                    .Where(t => t.PositiveTraitID.HasValue)
                    .Select(t => t.PositiveTraits.TraitName)
                    .Take(3)
                    .ToList();

                var negative = traits
                    .Where(t => t.NegativeTraitID.HasValue)
                    .Select(t => t.NegativeTraits.TraitName)
                    .Take(2)
                    .ToList();

                string result = "Характер спокойный, уравновешенный. ";

                if (positive.Any())
                    result += $"Положительные качества: {string.Join(", ", positive)}. ";

                if (negative.Any())
                    result += $"Требует внимания: {string.Join(", ", negative)}.";

                return result;
            }
        }
        #endregion

        // =====================================================================
        #region ОБРАБОТЧИКИ МЕНЮ
        // =====================================================================
        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void Registration_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
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
        }

        private void StudentCards_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null && menuItem.Header.ToString() == "Все студенты")
            {
                if (GroupComboBox.Items.Count > 0)
                {
                    GroupComboBox.SelectedIndex = 0;
                }

                StudentsDataGrid.ItemsSource = allStudents;
                StatusTextBlock.Text = $"Статус: Показаны все студенты ({allStudents.Count} чел.)";
                FooterStatusText.Text = $"Все студенты ({allStudents.Count} записей)";
                CurrentGroupText.Text = "Группа: Все группы";
                UpdateRecordCount();
            }
        }

        private void Groups_Click(object sender, RoutedEventArgs e)
        {
            if (UserRole != "Администратор")
            {
                MessageBox.Show("У вас нет прав для просмотра групп", "Доступ запрещен",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                double currentLeft = this.Left;
                double currentTop = this.Top;
                double currentWidth = this.Width;
                double currentHeight = this.Height;

                this.Hide();

                GroupsWindow groupsWindow = new GroupsWindow
                {
                    Left = currentLeft,
                    Top = currentTop,
                    Width = currentWidth,
                    Height = currentHeight,
                    WindowStartupLocation = WindowStartupLocation.Manual
                };

                groupsWindow.Closed += (s, args) =>
                {
                    this.Show();
                    LoadGroups();
                    LoadStudents();
                };

                groupsWindow.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Журнал группы колледжа\nВерсия 1.0\n\n© 2026",
                          "О программе", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        #endregion

        // =====================================================================
        #region ЗАГЛУШКИ ДЛЯ НЕРЕАЛИЗОВАННЫХ ФУНКЦИЙ
        // =====================================================================
        private void GenerateSocialPassport_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        private void GenerateCharacteristic_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        private void GeneratePortfolio_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        private void ExportToExcel_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        private void ExportToWord_Click(object sender, RoutedEventArgs e) => ShowNotImplemented();
        // ... остальные заглушки
        #endregion

        // =====================================================================
        #region МЕРОПРИЯТИЯ
        // =====================================================================
        private void EventsAdmin_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var eventsWindow = new EventsWindow(isAdminMode: true);
                eventsWindow.Owner = this;
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
                var eventsWindow = new EventsWindow(isAdminMode: false, groupId: UserGroupId);
                eventsWindow.Owner = this;
                eventsWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        #endregion
    }
}