using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace Kursach.APages
{
    public partial class GroupsWindow : Window
    {
        private vsstuEntities db = new vsstuEntities();
        private List<GroupViewModel> allGroups = new List<GroupViewModel>();
        private List<GroupViewModel> filteredGroups = new List<GroupViewModel>();
        public string CurrentDate => DateTime.Now.ToString("dd.MM.yyyy");

        public GroupsWindow()
        {
            InitializeComponent();
            DataContext = this;
            LoadGroups();
            LoadFacultiesToFilter();
        }

       
        private void LoadFacultiesToFilter()
        {
            try
            {
                var faculties = db.Faculties
                    .OrderBy(f => f.FacultyName)
                    .Select(f => f.FacultyName)
                    .ToList();

                FacultyFilterComboBox.Items.Clear();
                FacultyFilterComboBox.Items.Add(new ComboBoxItem { Content = "Все факультеты", IsSelected = true });

                foreach (var faculty in faculties)
                {
                    FacultyFilterComboBox.Items.Add(new ComboBoxItem { Content = faculty });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки факультетов: {ex.Message}");
            }
        }

        private void LoadGroups()
        {
            try
            {
                allGroups = (from g in db.Groups
                             join s in db.Specialties on g.SpecialtyID equals s.SpecialtyID into sj
                             from specialty in sj.DefaultIfEmpty()
                             join f in db.Faculties on specialty.FacultyID equals f.FacultyID into fj
                             from faculty in fj.DefaultIfEmpty()
                             select new GroupViewModel
                             {
                                 GroupID = g.GroupID,
                                 GroupName = g.GroupName,
                                 Course = g.Course,
                                 SpecialtyID = g.SpecialtyID,
                                 SpecialtyName = specialty != null ? specialty.SpecialtyName : "Не указана",
                                 SpecialtyCode = specialty != null ? specialty.SpecialtyCode : "",
                                 FacultyName = faculty != null ? faculty.FacultyName : "Не указан",
                                 StudentCount = g.Students.Count(s => s.IsActive == true),
                                 FormOfEducation = g.FormOfEducation,
                                 AcademicYear = g.AcademicYear,
                                 Language = g.Language,
                                 CreatedAt = g.CreatedAt
                             }).OrderBy(g => g.GroupName).ToList();

                allGroups = allGroups.OrderBy(g => g.GroupName).ToList();
                filteredGroups = new List<GroupViewModel>(allGroups);
                ApplyFilters();


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus("Статус: Ошибка загрузки", "#E74C3C");
            }
        }

        private void ApplyFilters()
        {
            try
            {
                var filtered = allGroups.AsEnumerable();

                var courseItem = CourseFilterComboBox.SelectedItem as ComboBoxItem;
                if (courseItem != null && courseItem.Content.ToString() != "Все курсы")
                {
                    int course = int.Parse(courseItem.Content.ToString().Replace(" курс", ""));
                    filtered = filtered.Where(g => g.Course == course);
                }

                var facultyItem = FacultyFilterComboBox.SelectedItem as ComboBoxItem;
                if (facultyItem != null && facultyItem.Content.ToString() != "Все факультеты")
                {
                    string faculty = facultyItem.Content.ToString();
                    filtered = filtered.Where(g => g.FacultyName == faculty);
                }

                if (!string.IsNullOrWhiteSpace(SearchTextBox?.Text))
                {
                    string search = SearchTextBox.Text.ToLower();
                    filtered = filtered.Where(g =>
                        (g.GroupName != null && g.GroupName.ToLower().Contains(search)) ||
                        (g.SpecialtyName != null && g.SpecialtyName.ToLower().Contains(search)) ||
                        (g.SpecialtyCode != null && g.SpecialtyCode.ToLower().Contains(search)) ||
                        (g.FacultyName != null && g.FacultyName.ToLower().Contains(search))
                    );
                }

                filteredGroups = filtered.ToList();
                GroupsDataGrid.ItemsSource = filteredGroups;

                StatusTextBlock.Text = $"Статус: Показано {filteredGroups.Count} групп из {allGroups.Count}";
                RecordCountText.Text = $"Всего записей: {filteredGroups.Count}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка фильтрации: {ex.Message}");
            }
        }

        private void UpdateStatus(string text, string colorHex)
        {
            StatusTextBlock.Text = text;
            StatusTextBlock.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
            StatusIndicator.Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorHex));
        }

        private void LoadDataButton_Click(object sender, RoutedEventArgs e)
        {
            LoadGroups();
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void CourseFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void FacultyFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void AllGroups_Click(object sender, RoutedEventArgs e)
        {
            CourseFilterComboBox.SelectedIndex = 0;
            FacultyFilterComboBox.SelectedIndex = 0;
            SearchTextBox.Text = "";
        }

        private void FilterByCourse_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                string header = menuItem.Header.ToString();
                for (int i = 0; i < CourseFilterComboBox.Items.Count; i++)
                {
                    var item = CourseFilterComboBox.Items[i] as ComboBoxItem;
                    if (item.Content.ToString() == header)
                    {
                        CourseFilterComboBox.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void FilterByFaculty_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                string header = menuItem.Header.ToString();
                for (int i = 0; i < FacultyFilterComboBox.Items.Count; i++)
                {
                    var item = FacultyFilterComboBox.Items[i] as ComboBoxItem;
                    if (item.Content.ToString() == header)
                    {
                        FacultyFilterComboBox.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new ADialogs.AddEditGroupDialog();
            dialog.Owner = this;
            dialog.Title = "Добавление группы";

            if (dialog.ShowDialog() == true)
            {
                LoadGroups();
            }
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите группу для редактирования", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selected = GroupsDataGrid.SelectedItem as GroupViewModel;
            var dialog = new ADialogs.AddEditGroupDialog(selected.GroupID);
            dialog.Owner = this;
            dialog.Title = "Редактирование группы";

            if (dialog.ShowDialog() == true)
            {
                LoadGroups();
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите группу для удаления", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selected = GroupsDataGrid.SelectedItem as GroupViewModel;

            if (selected.StudentCount > 0)
            {
                MessageBox.Show($"Невозможно удалить группу {selected.GroupName}, так как в ней числятся {selected.StudentCount} студентов.",
                              "Ошибка удаления", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show($"Удалить группу {selected.GroupName}?", "Подтверждение удаления",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    var group = db.Groups.Find(selected.GroupID);
                    if (group != null)
                    {
                        db.Groups.Remove(group);
                        db.SaveChanges();
                        LoadGroups();
                        UpdateStatus("Статус: Группа удалена", "#27AE60");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void GroupsDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem != null)
            {
                EditButton_Click(sender, null);
            }
        }

        private void Search_Click(object sender, RoutedEventArgs e)
        {
            SearchTextBox.Focus();
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.FileName = $"Группы_{DateTime.Now:yyyyMMdd}.xlsx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    MessageBox.Show($"Группы экспортированы в файл:\n{saveFileDialog.FileName}", "Экспорт завершен",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word документы (*.docx)|*.docx|Все файлы (*.*)|*.*";
                saveFileDialog.DefaultExt = "docx";
                saveFileDialog.FileName = $"Группы_{DateTime.Now:yyyyMMdd}.docx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    MessageBox.Show($"Документ сохранен:\n{saveFileDialog.FileName}", "Экспорт завершен",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CloseWindow_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void GroupsListReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word документы (*.docx)|*.docx";
                saveFileDialog.FileName = $"Список_групп_{DateTime.Now:yyyyMMdd}.docx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    MessageBox.Show("Отчет 'Список групп' создан", "Успех",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания отчета: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GroupsStatistics_Click(object sender, RoutedEventArgs e)
        {
            int totalGroups = allGroups.Count;
            int totalStudents = db.Students.Count(s => s.IsActive == true);
            int avgStudentsPerGroup = totalGroups > 0 ? totalStudents / totalGroups : 0;

            string stats = $"Всего групп: {totalGroups}\n" +
                          $"Всего студентов: {totalStudents}\n" +
                          $"Среднее студентов в группе: {avgStudentsPerGroup}\n\n" +
                          $"Распределение по курсам:\n";

            for (int i = 1; i <= 4; i++)
            {
                int count = allGroups.Count(g => g.Course == i);
                stats += $"{i} курс: {count} групп\n";
            }

            MessageBox.Show(stats, "Статистика по группам",
                          MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GroupsByCourseReport_Click(object sender, RoutedEventArgs e)
        {
            string report = "РАСПРЕДЕЛЕНИЕ ГРУПП ПО КУРСАМ\n\n";
            for (int i = 1; i <= 4; i++)
            {
                var groups = allGroups.Where(g => g.Course == i).ToList();
                report += $"{i} КУРС ({groups.Count} групп):\n";
                foreach (var group in groups)
                {
                    report += $"  - {group.GroupName} ({group.StudentCount} студ.)\n";
                }
                report += "\n";
            }

            MessageBox.Show(report, "Распределение по курсам",
                          MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GroupsByFacultyReport_Click(object sender, RoutedEventArgs e)
        {
            var faculties = allGroups.GroupBy(g => g.FacultyName)
                                     .Select(g => new { Faculty = g.Key, Count = g.Count() })
                                     .OrderByDescending(g => g.Count)
                                     .ToList();

            string report = "РАСПРЕДЕЛЕНИЕ ГРУПП ПО ФАКУЛЬТЕТАМ\n\n";
            foreach (var faculty in faculties)
            {
                report += $"{faculty.Faculty}: {faculty.Count} групп\n";
            }

            MessageBox.Show(report, "Распределение по факультетам",
                          MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GroupStudents_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите группу для просмотра студентов", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selected = GroupsDataGrid.SelectedItem as GroupViewModel;

            var students = db.Students
                .Where(s => s.GroupID == selected.GroupID && s.IsActive == true)
                .Select(s => new
                {
                    s.LastName,
                    s.FirstName,
                    s.MiddleName,
                    s.Phone,
                    s.Email
                })
                .ToList();

            string studentList = $"СТУДЕНТЫ ГРУППЫ {selected.GroupName}\n\n";
            if (students.Any())
            {
                foreach (var student in students)
                {
                    studentList += $"{student.LastName} {student.FirstName} {student.MiddleName}\n";
                    studentList += $"  Тел: {student.Phone ?? "не указан"}\n";
                    studentList += $"  Email: {student.Email ?? "не указан"}\n\n";
                }
            }
            else
            {
                studentList += "В группе нет активных студентов";
            }

            MessageBox.Show(studentList, $"Студенты группы {selected.GroupName}",
                          MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GroupStatistics_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите группу", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selected = GroupsDataGrid.SelectedItem as GroupViewModel;

            var students = db.Students.Where(s => s.GroupID == selected.GroupID).ToList();
            int activeStudents = students.Count(s => s.IsActive == true);
            int inactiveStudents = students.Count(s => s.IsActive == false);

            int male = students.Count(s => s.Gender == "М");
            int female = students.Count(s => s.Gender == "Ж");

            int orphans = students.Count(s => s.IsOrphan == true);
            int disabled = students.Count(s => s.IsDisabled == true);
            int largeFamily = students.Count(s => s.IsFromLargeFamily == true);
            int lowIncome = students.Count(s => s.IsLowIncome == true);
            int employed = students.Count(s => s.IsEmployed == true);

            string stats = $"СТАТИСТИКА ГРУППЫ {selected.GroupName}\n\n" +
                          $"Всего студентов: {students.Count}\n" +
                          $"  Активных: {activeStudents}\n" +
                          $"  Неактивных: {inactiveStudents}\n\n" +
                          $"По полу:\n" +
                          $"  Юноши: {male}\n" +
                          $"  Девушки: {female}\n\n" +
                          $"Социальный статус:\n" +
                          $"  Сироты: {orphans}\n" +
                          $"  Инвалиды: {disabled}\n" +
                          $"  Многодетные: {largeFamily}\n" +
                          $"  Малоимущие: {lowIncome}\n" +
                          $"  Работающие: {employed}";

            MessageBox.Show(stats, $"Статистика группы {selected.GroupName}",
                          MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void HelpContents_Click(object sender, RoutedEventArgs e)
        {
            string help = "РАБОТА С ГРУППАМИ\n\n" +
                         "Основные действия:\n" +
                         "- Добавление группы: кнопка 'Добавить' или меню\n" +
                         "- Редактирование: выбрать группу и нажать 'Редактировать' или двойной клик\n" +
                         "- Удаление: выбрать группу и нажать 'Удалить'\n\n" +
                         "Фильтрация:\n" +
                         "- По курсу: выберите курс в верхней панели\n" +
                         "- По факультету: выберите факультет в верхней панели\n" +
                         "- Поиск: введите название группы в поле поиска\n\n" +
                         "Отчеты:\n" +
                         "- Список групп - полный список всех групп\n" +
                         "- Статистика по группам - общая статистика\n" +
                         "- Распределение по курсам/факультетам";

            MessageBox.Show(help, "Справка", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Управление группами\nВерсия 1.0\n\n© 2026\n\nРазработано для колледжа",
                          "О программе", MessageBoxButton.OK, MessageBoxImage.Information);
        }

      
        public class GroupViewModel
        {
            public int GroupID { get; set; }
            public string GroupName { get; set; }
            public int? Course { get; set; }
            public int? SpecialtyID { get; set; }
            public string SpecialtyName { get; set; }
            public string SpecialtyCode { get; set; }
            public string FacultyName { get; set; }
            public int StudentCount { get; set; }
            public string FormOfEducation { get; set; }
            public string AcademicYear { get; set; }
            public string Language { get; set; }
            public DateTime? CreatedAt { get; set; }
        }
    }
}