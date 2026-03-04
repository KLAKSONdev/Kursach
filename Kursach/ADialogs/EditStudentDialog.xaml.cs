using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Kursach.ADialogs
{
    public partial class EditStudentDialog : Window
    {
        private vsstuEntities db = new vsstuEntities();
        private int _studentId;

        public EditStudentDialog(int studentId)
        {
            InitializeComponent();
            _studentId = studentId;
            LoadGroups();
            LoadStudentData();

            // Обработчик изменения даты рождения для автоматического расчета возраста
            BirthDatePicker.SelectedDateChanged += BirthDatePicker_SelectedDateChanged;
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

        private void LoadStudentData()
        {
            try
            {
                var student = db.Students.Find(_studentId);
                if (student == null)
                {
                    MessageBox.Show("Студент не найден", "Ошибка");
                    Close();
                    return;
                }

                // Основная информация
                LastNameTextBox.Text = student.LastName;
                FirstNameTextBox.Text = student.FirstName;
                MiddleNameTextBox.Text = student.MiddleName ?? "";
                GroupComboBox.SelectedValue = student.GroupID;

                // Пол
                if (student.Gender == "М")
                    GenderComboBox.SelectedIndex = 0;
                else if (student.Gender == "Ж")
                    GenderComboBox.SelectedIndex = 1;

                BirthDatePicker.SelectedDate = student.BirthDate;
                UpdateAge();

                // Документы
                StudentCardNumberTextBox.Text = student.StudentCardNumber ?? "";
                PersonalNumberTextBox.Text = student.PersonalNumber ?? "";
                EducationBeforeTextBox.Text = student.EducationBefore ?? "";
                EducationDocumentTextBox.Text = student.EducationDocument ?? "";

                // Личные данные
                BirthPlaceTextBox.Text = student.BirthPlace ?? "";
                NationalityTextBox.Text = student.Nationality ?? "";
                CitizenshipTextBox.Text = student.Citizenship ?? "";

                // Контакты
                PhoneTextBox.Text = student.Phone ?? "";
                EmailTextBox.Text = student.Email ?? "";
                ParentsPhoneTextBox.Text = student.ParentsPhone ?? "";

                // Адреса
                RegistrationAddressTextBox.Text = student.RegistrationAddress ?? "";
                ResidentialAddressTextBox.Text = student.ResidentialAddress ?? "";

                // Социальный статус
                IsOrphanCheckBox.IsChecked = student.IsOrphan ?? false;
                IsDisabledCheckBox.IsChecked = student.IsDisabled ?? false;
                IsFromLargeFamilyCheckBox.IsChecked = student.IsFromLargeFamily ?? false;
                IsLowIncomeCheckBox.IsChecked = student.IsLowIncome ?? false;

                // Трудоустройство
                IsEmployedCheckBox.IsChecked = student.IsEmployed ?? false;
                WorkPlaceTextBox.Text = student.WorkPlace ?? "";
                WorkPositionTextBox.Text = student.WorkPosition ?? "";

                // Даты и статус
                EnrollmentDatePicker.SelectedDate = student.EnrollmentDate;
                GraduationDatePicker.SelectedDate = student.GraduationDate;

                if (student.IsActive == true)
                    StatusComboBox.SelectedIndex = 0;
                else
                    StatusComboBox.SelectedIndex = 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}", "Ошибка");
            }
        }

        private void UpdateAge()
        {
            if (BirthDatePicker.SelectedDate.HasValue)
            {
                var today = DateTime.Today;
                var birthDate = BirthDatePicker.SelectedDate.Value;
                int age = today.Year - birthDate.Year;

                if (birthDate > today.AddYears(-age))
                    age--;

                AgeTextBox.Text = age.ToString();
            }
            else
            {
                AgeTextBox.Text = "";
            }
        }

        private void BirthDatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateAge();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка обязательных полей
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
                var student = db.Students.Find(_studentId);
                if (student == null)
                {
                    MessageBox.Show("Студент не найден", "Ошибка");
                    Close();
                    return;
                }

                // Основная информация
                student.LastName = LastNameTextBox.Text.Trim();
                student.FirstName = FirstNameTextBox.Text.Trim();
                student.MiddleName = string.IsNullOrWhiteSpace(MiddleNameTextBox.Text) ? null : MiddleNameTextBox.Text.Trim();
                student.GroupID = GroupComboBox.SelectedValue as int?;
                student.Gender = (GenderComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                student.BirthDate = BirthDatePicker.SelectedDate;

                // Документы
                student.StudentCardNumber = string.IsNullOrWhiteSpace(StudentCardNumberTextBox.Text) ? null : StudentCardNumberTextBox.Text.Trim();
                student.PersonalNumber = string.IsNullOrWhiteSpace(PersonalNumberTextBox.Text) ? null : PersonalNumberTextBox.Text.Trim();
                student.EducationBefore = string.IsNullOrWhiteSpace(EducationBeforeTextBox.Text) ? null : EducationBeforeTextBox.Text.Trim();
                student.EducationDocument = string.IsNullOrWhiteSpace(EducationDocumentTextBox.Text) ? null : EducationDocumentTextBox.Text.Trim();

                // Личные данные
                student.BirthPlace = string.IsNullOrWhiteSpace(BirthPlaceTextBox.Text) ? null : BirthPlaceTextBox.Text.Trim();
                student.Nationality = string.IsNullOrWhiteSpace(NationalityTextBox.Text) ? null : NationalityTextBox.Text.Trim();
                student.Citizenship = string.IsNullOrWhiteSpace(CitizenshipTextBox.Text) ? null : CitizenshipTextBox.Text.Trim();

                // Контакты
                student.Phone = string.IsNullOrWhiteSpace(PhoneTextBox.Text) ? null : PhoneTextBox.Text.Trim();
                student.Email = string.IsNullOrWhiteSpace(EmailTextBox.Text) ? null : EmailTextBox.Text.Trim();
                student.ParentsPhone = string.IsNullOrWhiteSpace(ParentsPhoneTextBox.Text) ? null : ParentsPhoneTextBox.Text.Trim();

                // Адреса
                student.RegistrationAddress = string.IsNullOrWhiteSpace(RegistrationAddressTextBox.Text) ? null : RegistrationAddressTextBox.Text.Trim();
                student.ResidentialAddress = string.IsNullOrWhiteSpace(ResidentialAddressTextBox.Text) ? null : ResidentialAddressTextBox.Text.Trim();

                // Социальный статус
                student.IsOrphan = IsOrphanCheckBox.IsChecked;
                student.IsDisabled = IsDisabledCheckBox.IsChecked;
                student.IsFromLargeFamily = IsFromLargeFamilyCheckBox.IsChecked;
                student.IsLowIncome = IsLowIncomeCheckBox.IsChecked;

                // Трудоустройство
                student.IsEmployed = IsEmployedCheckBox.IsChecked;
                student.WorkPlace = IsEmployedCheckBox.IsChecked == true ?
                    (string.IsNullOrWhiteSpace(WorkPlaceTextBox.Text) ? null : WorkPlaceTextBox.Text.Trim()) : null;
                student.WorkPosition = IsEmployedCheckBox.IsChecked == true ?
                    (string.IsNullOrWhiteSpace(WorkPositionTextBox.Text) ? null : WorkPositionTextBox.Text.Trim()) : null;

                // Даты и статус
                student.EnrollmentDate = EnrollmentDatePicker.SelectedDate;
                student.GraduationDate = GraduationDatePicker.SelectedDate;
                student.IsActive = StatusComboBox.SelectedIndex == 0;
                student.UpdatedAt = DateTime.Now;

                db.SaveChanges();

                MessageBox.Show("Данные студента успешно обновлены!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                string error = "Ошибка при сохранении:\n";
                if (ex.InnerException != null)
                    error += ex.InnerException.Message;
                else
                    error += ex.Message;

                MessageBox.Show(error, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}