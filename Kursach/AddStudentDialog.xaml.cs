using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Kursach.Dialogs
{
    public partial class AddStudentDialog : Window
    {
        private vsstuEntities db = new vsstuEntities();

        public AddStudentDialog()
        {
            InitializeComponent();
            LoadGroups();
        }

        private void LoadGroups()
        {
            try
            {
                var groups = db.Groups
                    .OrderBy(g => g.GroupName)
                    .Select(g => new { g.GroupID, g.GroupName })
                    .ToList();

                GroupComboBox.ItemsSource = groups;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}", "Ошибка");
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка только самых необходимых полей
            if (string.IsNullOrWhiteSpace(LastNameTextBox.Text))
            {
                MessageBox.Show("Введите фамилию", "Ошибка");
                LastNameTextBox.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(FirstNameTextBox.Text))
            {
                MessageBox.Show("Введите имя", "Ошибка");
                FirstNameTextBox.Focus();
                return;
            }

            try
            {
                var student = new Students
                {
                    // Только основные поля
                    LastName = LastNameTextBox.Text.Trim(),
                    FirstName = FirstNameTextBox.Text.Trim(),
                    MiddleName = string.IsNullOrWhiteSpace(MiddleNameTextBox.Text) ? null : MiddleNameTextBox.Text.Trim(),
                    GroupID = GroupComboBox.SelectedValue as int?,
                    Gender = (GenderComboBox.SelectedItem as ComboBoxItem)?.Content.ToString(),
                    BirthDate = BirthDatePicker.SelectedDate,
                    Phone = string.IsNullOrWhiteSpace(PhoneTextBox.Text) ? null : PhoneTextBox.Text.Trim(),
                    Email = string.IsNullOrWhiteSpace(EmailTextBox.Text) ? null : EmailTextBox.Text.Trim(),

                    // Соц статус
                    IsOrphan = IsOrphanCheckBox.IsChecked,
                    IsDisabled = IsDisabledCheckBox.IsChecked,
                    IsFromLargeFamily = IsFromLargeFamilyCheckBox.IsChecked,
                    IsLowIncome = IsLowIncomeCheckBox.IsChecked,

                    // Системные поля
                    IsActive = true,
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now

                    // Все остальные поля будут NULL - теперь это разрешено!
                };

                db.Students.Add(student);
                db.SaveChanges();

                MessageBox.Show("Студент успешно добавлен!", "Успех");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}