using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Kursach.ADialogs
{
    public partial class AddEditGroupDialog : Window
    {
        private vsstuEntities db = new vsstuEntities();
        private int? groupId = null;

        // Конструктор для добавления
        public AddEditGroupDialog()
        {
            InitializeComponent();
            LoadSpecialties();
            TitleText.Text = "Добавление группы";
        }

        // Конструктор для редактирования
        public AddEditGroupDialog(int id)
        {
            InitializeComponent();
            this.groupId = id;
            LoadSpecialties();
            LoadGroupData();
            TitleText.Text = "Редактирование группы";
        }

        // Загрузка специальностей
        private void LoadSpecialties()
        {
            var specialties = db.Specialties
                .Select(s => new
                {
                    s.SpecialtyID,
                    SpecialtyName = s.SpecialtyCode + " " + s.SpecialtyName
                })
                .ToList();

            SpecialtyComboBox.ItemsSource = specialties;
        }

        // Загрузка данных группы для редактирования
        private void LoadGroupData()
        {
            if (groupId == null) return;

            var group = db.Groups.Find(groupId);
            if (group != null)
            {
                GroupNameTextBox.Text = group.GroupName;

                if (group.Course.HasValue)
                    CourseComboBox.SelectedIndex = group.Course.Value - 1;

                if (group.SpecialtyID.HasValue)
                    SpecialtyComboBox.SelectedValue = group.SpecialtyID.Value;

                if (!string.IsNullOrEmpty(group.FormOfEducation))
                {
                    for (int i = 0; i < FormComboBox.Items.Count; i++)
                    {
                        var item = FormComboBox.Items[i] as ComboBoxItem;
                        if (item.Content.ToString() == group.FormOfEducation)
                        {
                            FormComboBox.SelectedIndex = i;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(group.AcademicYear))
                {
                    for (int i = 0; i < YearComboBox.Items.Count; i++)
                    {
                        var item = YearComboBox.Items[i] as ComboBoxItem;
                        if (item.Content.ToString() == group.AcademicYear)
                        {
                            YearComboBox.SelectedIndex = i;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(group.Language))
                {
                    for (int i = 0; i < LanguageComboBox.Items.Count; i++)
                    {
                        var item = LanguageComboBox.Items[i] as ComboBoxItem;
                        if (item.Content.ToString() == group.Language)
                        {
                            LanguageComboBox.SelectedIndex = i;
                            break;
                        }
                    }
                }
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка заполнения обязательных полей
            if (string.IsNullOrWhiteSpace(GroupNameTextBox.Text))
            {
                MessageBox.Show("Введите название группы", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                GroupNameTextBox.Focus();
                return;
            }

            if (SpecialtyComboBox.SelectedValue == null)
            {
                MessageBox.Show("Выберите специальность", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                SpecialtyComboBox.Focus();
                return;
            }

            try
            {
                if (groupId == null) // Добавление
                {
                    // Проверка на дубликат названия
                    bool exists = db.Groups.Any(g => g.GroupName == GroupNameTextBox.Text.Trim());
                    if (exists)
                    {
                        MessageBox.Show("Группа с таким названием уже существует", "Ошибка",
                                      MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var group = new Groups
                    {
                        GroupName = GroupNameTextBox.Text.Trim(),
                        Course = CourseComboBox.SelectedIndex + 1,
                        SpecialtyID = (int)SpecialtyComboBox.SelectedValue,
                        FormOfEducation = (FormComboBox.SelectedItem as ComboBoxItem)?.Content.ToString(),
                        AcademicYear = (YearComboBox.SelectedItem as ComboBoxItem)?.Content.ToString(),
                        Language = (LanguageComboBox.SelectedItem as ComboBoxItem)?.Content.ToString(),
                        CreatedAt = DateTime.Now
                    };

                    db.Groups.Add(group);
                    db.SaveChanges();

                    MessageBox.Show("Группа успешно добавлена", "Успех",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else // Редактирование
                {
                    var group = db.Groups.Find(groupId);
                    if (group != null)
                    {
                        // Проверка на дубликат названия (исключая текущую группу)
                        bool exists = db.Groups.Any(g => g.GroupName == GroupNameTextBox.Text.Trim()
                                                      && g.GroupID != groupId);
                        if (exists)
                        {
                            MessageBox.Show("Группа с таким названием уже существует", "Ошибка",
                                          MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        group.GroupName = GroupNameTextBox.Text.Trim();
                        group.Course = CourseComboBox.SelectedIndex + 1;
                        group.SpecialtyID = (int)SpecialtyComboBox.SelectedValue;
                        group.FormOfEducation = (FormComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                        group.AcademicYear = (YearComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                        group.Language = (LanguageComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();

                        db.SaveChanges();

                        MessageBox.Show("Группа успешно обновлена", "Успех",
                                      MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}