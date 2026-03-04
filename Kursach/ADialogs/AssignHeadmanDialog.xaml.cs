using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Kursach.ADialogs
{
    public partial class AssignHeadmanDialog : Window
    {
        private vsstuEntities db = new vsstuEntities();
        private int _groupId;
        private int? _currentHeadmanId;

        public AssignHeadmanDialog(int groupId)
        {
            InitializeComponent();
            _groupId = groupId;
            LoadGroupInfo();
            LoadStudents();
        }

        private void LoadGroupInfo()
        {
            try
            {
                var group = db.Groups.Find(_groupId);
                if (group != null)
                {
                    GroupNameText.Text = group.GroupName;

                    // Находим текущего старосту
                    var currentHeadman = db.Students
                        .FirstOrDefault(s => s.GroupID == _groupId && s.IsHeadman == true);

                    if (currentHeadman != null)
                    {
                        _currentHeadmanId = currentHeadman.StudentID;
                        CurrentHeadmanText.Text = $"Текущий староста: {currentHeadman.LastName} {currentHeadman.FirstName} {currentHeadman.MiddleName}".Trim();
                    }
                    else
                    {
                        CurrentHeadmanText.Text = "Текущий староста: не назначен";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки информации о группе: {ex.Message}", "Ошибка");
            }
        }

        private void LoadStudents()
        {
            try
            {
                var students = db.Students
                    .Where(s => s.GroupID == _groupId && s.IsActive == true)
                    .OrderBy(s => s.LastName)
                    .ThenBy(s => s.FirstName)
                    .Select(s => new
                    {
                        s.StudentID,
                        DisplayName = $"{s.LastName} {s.FirstName} {s.MiddleName}".Trim() + (s.IsHeadman == true ? " (староста)" : ""),
                        IsHeadman = s.IsHeadman ?? false
                    })
                    .ToList();

                StudentsListBox.ItemsSource = students;
                StudentsListBox.DisplayMemberPath = "DisplayName";
                StudentsListBox.SelectedValuePath = "StudentID";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки студентов: {ex.Message}", "Ошибка");
            }
        }

        private void AssignButton_Click(object sender, RoutedEventArgs e)
        {
            if (StudentsListBox.SelectedItem == null)
            {
                MessageBox.Show("Выберите студента", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var selectedStudentId = (int)StudentsListBox.SelectedValue;

                // Если выбран тот же студент, который уже староста
                if (selectedStudentId == _currentHeadmanId)
                {
                    MessageBox.Show("Этот студент уже является старостой", "Информация",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Снимаем статус старосты с текущего (если есть)
                if (_currentHeadmanId.HasValue)
                {
                    var oldHeadman = db.Students.Find(_currentHeadmanId.Value);
                    if (oldHeadman != null)
                    {
                        oldHeadman.IsHeadman = false;
                    }
                }

                // Назначаем нового старосту
                var newHeadman = db.Students.Find(selectedStudentId);
                if (newHeadman != null)
                {
                    newHeadman.IsHeadman = true;
                    db.SaveChanges();

                    MessageBox.Show($"Староста назначен: {newHeadman.LastName} {newHeadman.FirstName}",
                                  "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    DialogResult = true;
                    Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при назначении старосты: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_currentHeadmanId.HasValue)
            {
                MessageBox.Show("В группе нет назначенного старосты", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (MessageBox.Show("Снять старосту с должности?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    var currentHeadman = db.Students.Find(_currentHeadmanId.Value);
                    if (currentHeadman != null)
                    {
                        currentHeadman.IsHeadman = false;
                        db.SaveChanges();

                        MessageBox.Show("Староста снят с должности", "Успех",
                                      MessageBoxButton.OK, MessageBoxImage.Information);
                        DialogResult = true;
                        Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}