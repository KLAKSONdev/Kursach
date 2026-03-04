using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Kursach.AWindows
{
    public partial class ParticipationWindow : Window
    {
        private vsstuEntities db = new vsstuEntities();
        private int _eventId;
        private int? _groupId;
        private List<StudentParticipationModel> students = new List<StudentParticipationModel>();

        public class StudentParticipationModel
        {
            public int StudentID { get; set; }
            public string FullName { get; set; }
            public string GroupName { get; set; }
            public int? GroupID { get; set; }
            public bool IsParticipating { get; set; }
            public string Role { get; set; }
            public string Result { get; set; }
            public int? ParticipationID { get; set; }
        }

        public ParticipationWindow(int eventId, int? groupId = null)
        {
            InitializeComponent();
            _eventId = eventId;
            _groupId = groupId;

            LoadEventInfo();
            LoadStudents();
        }

        private void LoadEventInfo()
        {
            try
            {
                var ev = db.Events.Find(_eventId);
                if (ev != null)
                {
                    EventTitleText.Text = ev.EventName;
                    EventDateText.Text = ev.EventDate?.ToString("dd.MM.yyyy") ?? "-";
                    EventLocationText.Text = ev.Location ?? "-";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки информации: {ex.Message}");
            }
        }

        private void LoadStudents()
        {
            try
            {
                // Получаем студентов (если староста - только его группа)
                var studentQuery = db.Students.Where(s => s.IsActive == true);

                if (_groupId.HasValue)
                {
                    studentQuery = studentQuery.Where(s => s.GroupID == _groupId.Value);
                }

                var dbStudents = studentQuery.ToList();

                // Получаем текущих участников
                var participants = db.EventParticipation
                    .Where(p => p.EventID == _eventId)
                    .ToDictionary(p => p.StudentID);

                students = dbStudents.Select(s => new StudentParticipationModel
                {
                    StudentID = s.StudentID,
                    FullName = $"{s.LastName} {s.FirstName} {s.MiddleName}".Trim(),
                    GroupName = s.Groups?.GroupName ?? "Без группы",
                    GroupID = s.GroupID,
                    IsParticipating = participants.ContainsKey(s.StudentID),
                    Role = participants.ContainsKey(s.StudentID) ? participants[s.StudentID].Role : "",
                    Result = participants.ContainsKey(s.StudentID) ? participants[s.StudentID].Result : "",
                    ParticipationID = participants.ContainsKey(s.StudentID) ? participants[s.StudentID].ParticipationID : (int?)null
                }).OrderBy(s => s.FullName).ToList();

                StudentsGrid.ItemsSource = students;
                StudentsGrid.IsReadOnly = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки студентов: {ex.Message}");
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем текущих участников из БД
                var existingParticipations = db.EventParticipation
                    .Where(p => p.EventID == _eventId)
                    .ToDictionary(p => p.StudentID);

                foreach (var student in students)
                {
                    if (student.IsParticipating)
                    {
                        if (existingParticipations.ContainsKey(student.StudentID))
                        {
                            // Обновляем существующую запись
                            var part = existingParticipations[student.StudentID];
                            part.Role = student.Role;
                            part.Result = student.Result;
                        }
                        else
                        {
                            // Добавляем новую запись
                            var newPart = new EventParticipation
                            {
                                EventID = _eventId,
                                StudentID = student.StudentID,
                                Role = student.Role,
                                Result = student.Result,
                                ParticipationStatus = "Участвовал"
                            };
                            db.EventParticipation.Add(newPart);
                        }
                    }
                    else
                    {
                        // Если есть запись, но студент не участвует - удаляем
                        if (existingParticipations.ContainsKey(student.StudentID))
                        {
                            db.EventParticipation.Remove(existingParticipations[student.StudentID]);
                        }
                    }
                }

                db.SaveChanges();
                MessageBox.Show("Участники сохранены!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}