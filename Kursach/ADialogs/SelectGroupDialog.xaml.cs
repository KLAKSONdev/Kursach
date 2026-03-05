using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace Kursach.ADialogs
{
    public partial class SelectGroupDialog : Window
    {
        private vsstuEntities db = new vsstuEntities();

        public int? SelectedGroupId { get; private set; }
        public string SelectedGroupName { get; private set; }

        public SelectGroupDialog()
        {
            InitializeComponent();
            LoadGroups();
        }

        private void LoadGroups()
        {
            try
            {
                var groups = db.Groups
                    .Select(g => new
                    {
                        g.GroupID,
                        g.GroupName,
                        g.Course,
                        SpecialtyName = g.Specialties != null ? g.Specialties.SpecialtyName : "Не указана",
                        StudentCount = g.Students.Count(s => s.IsActive == true)
                    })
                    .OrderBy(g => g.GroupName)
                    .ToList();

                GroupsDataGrid.ItemsSource = groups;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}", "Ошибка");
            }
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите группу", "Информация",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            dynamic selected = GroupsDataGrid.SelectedItem;
            SelectedGroupId = selected.GroupID;
            SelectedGroupName = selected.GroupName;

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void GroupsDataGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem != null)
            {
                SelectButton_Click(sender, null);
            }
        }
    }
}