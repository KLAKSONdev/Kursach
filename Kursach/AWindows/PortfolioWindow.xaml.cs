using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using Microsoft.Win32;
using System.Data.Entity;

namespace Kursach.AWindows
{
    public partial class PortfolioWindow : Window
    {
        private int _studentId;
        private vsstuEntities db = new vsstuEntities();

        // Модели для данных
        public class GradeModel
        {
            public string Subject { get; set; }
            public int? Grade { get; set; }
            public int? Semester { get; set; }
            public DateTime? Date { get; set; }
            public string Teacher { get; set; }
        }

        public class AchievementModel
        {
            public string AchievementName { get; set; }
            public string AchievementType { get; set; }
            public DateTime? Date { get; set; }
            public string Level { get; set; }
            public string Place { get; set; }
        }

        public class EventModel
        {
            public string EventName { get; set; }
            public string EventType { get; set; }
            public DateTime? Date { get; set; }
            public string Role { get; set; }
            public string Result { get; set; }
        }

        public class ParentModel
        {
            public string FullName { get; set; }
            public string Relationship { get; set; }
            public string Phone { get; set; }
            public string WorkPlace { get; set; }
            public string WorkPosition { get; set; }
        }

        public PortfolioWindow(int studentId)
        {
            InitializeComponent();
            _studentId = studentId;
            Loaded += PortfolioWindow_Loaded;
        }

        private void PortfolioWindow_Loaded(object sender, RoutedEventArgs e)
        {
            LoadStudentData();
            LoadGrades();
            LoadAchievements();
            LoadEvents();
            LoadParents();
            LoadCharacteristic();
        }

        private void LoadStudentData()
        {
            try
            {
                var student = db.Students
                    .Include(s => s.Groups)
                    .Include(s => s.FamilyTypes)
                    .Include(s => s.HealthStatus)
                    .FirstOrDefault(s => s.StudentID == _studentId);

                if (student != null)
                {
                    // ФИО
                    StudentFullNameText.Text = $"{student.LastName} {student.FirstName} {student.MiddleName}".Trim();

                    // Группа и курс
                    if (student.Groups != null)
                    {
                        StudentGroupText.Text = student.Groups.GroupName;
                        StudentCourseText.Text = student.Groups.Course?.ToString() ?? "Не указан";
                    }
                    else
                    {
                        StudentGroupText.Text = "Не указана";
                        StudentCourseText.Text = "Не указан";
                    }

                    // Студенческий билет
                    StudentCardNumberText.Text = $"№ {student.StudentCardNumber ?? "Не указан"}";

                    // Личные данные
                    BirthDateText.Text = student.BirthDate?.ToString("dd.MM.yyyy") ?? "Не указана";
                    BirthPlaceText.Text = student.BirthPlace ?? "Не указано";
                    GenderText.Text = student.Gender ?? "Не указан";
                    NationalityText.Text = student.Nationality ?? "Не указана";
                    CitizenshipText.Text = student.Citizenship ?? "Не указано";
                    PhoneText.Text = student.Phone ?? "Не указан";
                    EmailText.Text = student.Email ?? "Не указан";
                    ParentsPhoneText.Text = student.ParentsPhone ?? "Не указан";

                    // Адреса
                    RegistrationAddressText.Text = student.RegistrationAddress ?? "Не указан";
                    ResidentialAddressText.Text = student.ResidentialAddress ?? "Не указан";

                    // Социальный статус
                    IsOrphanCheck.IsChecked = student.IsOrphan ?? false;
                    IsDisabledCheck.IsChecked = student.IsDisabled ?? false;
                    IsFromLargeFamilyCheck.IsChecked = student.IsFromLargeFamily ?? false;
                    IsLowIncomeCheck.IsChecked = student.IsLowIncome ?? false;
                    // После загрузки социального статуса или в соответствующем месте
                    IsHeadmanCheckBox.IsChecked = student.IsHeadman ?? false;

                    var statuses = new List<string>();
                    if (student.IsOrphan == true) statuses.Add("Сирота");
                    if (student.IsDisabled == true) statuses.Add("Инвалид");
                    if (student.IsFromLargeFamily == true) statuses.Add("Многодетная семья");
                    if (student.IsLowIncome == true) statuses.Add("Малообеспеченный");
                    SocialStatusText.Text = statuses.Count > 0 ? string.Join(", ", statuses) : "Не указан";

                    // Образование
                    EducationBeforeText.Text = student.EducationBefore ?? "Не указано";
                    EducationDocumentText.Text = student.EducationDocument ?? "Не указан";
                    EnrollmentDateText.Text = student.EnrollmentDate?.ToString("dd.MM.yyyy") ?? "Не указана";
                    GraduationDateText.Text = student.GraduationDate?.ToString("dd.MM.yyyy") ?? "Не указана";

                    // Трудоустройство
                    IsEmployedCheck.IsChecked = student.IsEmployed ?? false;
                    WorkPlaceText.Text = student.WorkPlace ?? "Не указано";
                    WorkPositionText.Text = student.WorkPosition ?? "Не указана";

                    // Системная информация
                    SystemInfoText.Text = $"ID: {student.StudentID} | " +
                        $"Создано: {student.CreatedAt?.ToString("dd.MM.yyyy")} | " +
                        $"Обновлено: {student.UpdatedAt?.ToString("dd.MM.yyyy")}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных студента: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadGrades()
        {
            try
            {
                var grades = db.AcademicPerformance
                    .Where(g => g.StudentID == _studentId)
                    .OrderBy(g => g.Date)
                    .Select(g => new GradeModel
                    {
                        Subject = g.SubjectName,
                        Grade = g.Grade,
                        Semester = g.Semester,
                        Date = g.Date,
                        Teacher = g.Teacher
                    })
                    .ToList();

                GradesDataGrid.ItemsSource = grades;

                if (grades.Any())
                {
                    var avg = grades.Where(g => g.Grade.HasValue).Average(g => g.Grade.Value);
                    AvgGradeText.Text = avg.ToString("F2");
                }
                else
                {
                    AvgGradeText.Text = "0";
                }

                GradesCountText.Text = grades.Count.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки оценок: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadAchievements()
        {
            try
            {
                var achievements = db.Achievements
                    .Include(a => a.AchievementTypes)
                    .Where(a => a.StudentID == _studentId)
                    .OrderByDescending(a => a.Date)
                    .Select(a => new AchievementModel
                    {
                        AchievementName = a.AchievementName,
                        AchievementType = a.AchievementTypes != null ? a.AchievementTypes.TypeName : "Не указан",
                        Date = a.Date,
                        Level = a.Level,
                        Place = a.Place
                    })
                    .ToList();

                AchievementsDataGrid.ItemsSource = achievements;
                AchievementsCountText.Text = achievements.Count.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки достижений: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadEvents()
        {
            try
            {
                var events = db.EventParticipation
                    .Include(ep => ep.Events)
                    .Include(ep => ep.Events.EventTypes)
                    .Where(ep => ep.StudentID == _studentId)
                    .OrderByDescending(ep => ep.Events.EventDate)
                    .Select(ep => new EventModel
                    {
                        EventName = ep.Events.EventName,
                        EventType = ep.Events.EventTypes != null ? ep.Events.EventTypes.EventTypeName : "Не указан",
                        Date = ep.Events.EventDate,
                        Role = ep.Role,
                        Result = ep.Result
                    })
                    .ToList();

                EventsDataGrid.ItemsSource = events;
                EventsCountText.Text = events.Count.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки мероприятий: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadParents()
        {
            try
            {
                var parents = db.StudentParents
                    .Include(sp => sp.Parents)
                    .Where(sp => sp.StudentID == _studentId)
                    .Select(sp => new ParentModel
                    {
                        FullName = sp.Parents.LastName + " " + sp.Parents.FirstName + " " + sp.Parents.MiddleName,
                        Relationship = sp.Parents.Relationship,
                        Phone = sp.Parents.Phone,
                        WorkPlace = sp.Parents.WorkPlace,
                        WorkPosition = sp.Parents.WorkPosition
                    })
                    .ToList();

                ParentsDataGrid.ItemsSource = parents;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки родителей: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ========== МЕТОДЫ ДЛЯ ХАРАКТЕРИСТИКИ (СОХРАНЕНИЕ В БД) ==========

        private void LoadCharacteristic()
        {
            try
            {
                var characteristic = db.StudentCharacteristics
                    .FirstOrDefault(c => c.StudentID == _studentId);

                if (characteristic != null)
                {
                    CharacteristicTextBox.Text = characteristic.CharacteristicText;
                }
                else
                {
                    CharacteristicTextBox.Text = GetDefaultCharacteristic();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки характеристики: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                CharacteristicTextBox.Text = GetDefaultCharacteristic();
            }
        }

        private string GetDefaultCharacteristic()
        {
            var student = db.Students
                .Include(s => s.Groups)
                .FirstOrDefault(s => s.StudentID == _studentId);

            if (student == null) return "";

            return 
                   $"ЗАКЛЮЧЕНИЕ:\n" +
                   $"Студент за время обучения проявил себя как ...\n\n" +
                   $"ЛИЧНОСТНЫЕ КАЧЕСТВА:\n" +
                   $"УЧАСТИЕ В ОБЩЕСТВЕННОЙ ЖИЗНИ:\n" +
                   $"\n\n" +
                   $"РЕКОМЕНДАЦИИ:\n";
        }

        private void SaveCharacteristic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var characteristic = db.StudentCharacteristics
                    .FirstOrDefault(c => c.StudentID == _studentId);

                if (characteristic == null)
                {
                    // Создаем новую характеристику
                    characteristic = new StudentCharacteristics
                    {
                        StudentID = _studentId,
                        CharacteristicText = CharacteristicTextBox.Text,
                        CreatedDate = DateTime.Now,
                        UpdatedDate = DateTime.Now
                    };
                    db.StudentCharacteristics.Add(characteristic);
                }
                else
                {
                    // Обновляем существующую
                    characteristic.CharacteristicText = CharacteristicTextBox.Text;
                    characteristic.UpdatedDate = DateTime.Now;
                }

                db.SaveChanges();

                MessageBox.Show("Характеристика успешно сохранена в базу данных!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения характеристики: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearCharacteristic_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить характеристику и восстановить шаблон?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                CharacteristicTextBox.Text = GetDefaultCharacteristic();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ExportToPdf_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "PDF файлы (*.pdf)|*.pdf";
                saveFileDialog.DefaultExt = "pdf";
                saveFileDialog.FileName = $"Портфолио_{StudentFullNameText.Text}_{DateTime.Now:yyyyMMdd}.pdf";

                if (saveFileDialog.ShowDialog() == true)
                {
                    ExportToPdf(saveFileDialog.FileName);
                    MessageBox.Show($"Портфолио успешно экспортировано!\nФайл сохранен: {saveFileDialog.FileName}",
                                  "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в PDF: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToPdf(string filePath)
        {
            Document document = new Document(PageSize.A4, 30, 30, 40, 40);

            try
            {
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                document.Open();

                // Шрифт с поддержкой кириллицы
                BaseFont baseFont;
                string arialPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");

                if (File.Exists(arialPath))
                {
                    baseFont = BaseFont.CreateFont(arialPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                }
                else
                {
                    baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                }

                Font titleFont = new Font(baseFont, 16, Font.BOLD, new BaseColor(0, 51, 102));
                Font headerFont = new Font(baseFont, 14, Font.BOLD, new BaseColor(0, 51, 102));
                Font normalFont = new Font(baseFont, 11, Font.NORMAL, BaseColor.BLACK);
                Font boldFont = new Font(baseFont, 11, Font.BOLD, BaseColor.BLACK);
                Font labelFont = new Font(baseFont, 10, Font.NORMAL, new BaseColor(89, 89, 89));
                Font smallFont = new Font(baseFont, 9, Font.NORMAL, new BaseColor(128, 128, 128));

                // Титульный лист
                AddTitlePage(document, titleFont, headerFont, normalFont, smallFont, boldFont, labelFont);

                // Личные данные
                document.NewPage();
                AddSectionHeader(document, "1. ЛИЧНЫЕ ДАННЫЕ", headerFont);
                AddPersonalDataTable(document, normalFont, boldFont, labelFont);

                // Образование
                document.NewPage();
                AddSectionHeader(document, "2. ОБРАЗОВАНИЕ", headerFont);
                AddEducationSection(document, normalFont, boldFont, labelFont);

                // Социальный статус
                AddSectionHeader(document, "3. СОЦИАЛЬНЫЙ СТАТУС", headerFont);
                AddSocialStatusSection(document, normalFont, boldFont);

                // Трудоустройство
                AddSectionHeader(document, "4. ТРУДОУСТРОЙСТВО", headerFont);
                AddEmploymentSection(document, normalFont, boldFont);

                // Успеваемость
                document.NewPage();
                AddSectionHeader(document, "5. УСПЕВАЕМОСТЬ", headerFont);
                AddGradesSection(document, normalFont, boldFont, smallFont);

                // Достижения
                AddSectionHeader(document, "6. ДОСТИЖЕНИЯ", headerFont);
                AddAchievementsSection(document, normalFont, smallFont);

                // Мероприятия
                AddSectionHeader(document, "7. УЧАСТИЕ В МЕРОПРИЯТИЯХ", headerFont);
                AddEventsSection(document, normalFont, smallFont);

                // Родители
                document.NewPage();
                AddSectionHeader(document, "8. РОДИТЕЛИ", headerFont);
                AddParentsSection(document, normalFont, boldFont);

                // Характеристика
                document.NewPage();
                AddSectionHeader(document, "9. ХАРАКТЕРИСТИКА СТУДЕНТА", headerFont);
                AddCharacteristicSection(document, normalFont);

                // Подписи
                AddFooterSection(document, smallFont, boldFont);

                document.Close();
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка при создании PDF: {ex.Message}");
            }
            finally
            {
                if (document.IsOpen())
                    document.Close();
            }
        }

        #region PDF методы

        private void AddTitlePage(Document document, Font titleFont, Font headerFont, Font normalFont,
            Font smallFont, Font boldFont, Font labelFont)
        {
            Paragraph logo = new Paragraph("🎓", new Font(Font.FontFamily.HELVETICA, 40));
            logo.Alignment = Element.ALIGN_CENTER;
            logo.SpacingAfter = 10;
            document.Add(logo);

            Paragraph ministry = new Paragraph("МИНИСТЕРСТВО ОБРАЗОВАНИЯ", headerFont);
            ministry.Alignment = Element.ALIGN_CENTER;
            ministry.SpacingAfter = 5;
            document.Add(ministry);

            Paragraph college = new Paragraph("ПОРТФОЛИО СТУДЕНТА", new Font(headerFont.BaseFont, 20, Font.BOLD));
            college.Alignment = Element.ALIGN_CENTER;
            college.SpacingAfter = 40;
            document.Add(college);

            Paragraph studentName = new Paragraph(StudentFullNameText.Text.ToUpper(),
                new Font(titleFont.BaseFont, 18, Font.BOLD, new BaseColor(0, 51, 102)));
            studentName.Alignment = Element.ALIGN_CENTER;
            studentName.SpacingAfter = 10;
            document.Add(studentName);

            Paragraph groupInfo = new Paragraph($"Группа: {StudentGroupText.Text}     Курс: {StudentCourseText.Text}",
                new Font(normalFont.BaseFont, 14, Font.NORMAL));
            groupInfo.Alignment = Element.ALIGN_CENTER;
            groupInfo.SpacingAfter = 60;
            document.Add(groupInfo);

            PdfPTable infoTable = new PdfPTable(2);
            infoTable.WidthPercentage = 60;
            infoTable.HorizontalAlignment = Element.ALIGN_CENTER;
            infoTable.SpacingAfter = 40;

            AddPdfRow(infoTable, "Номер студенческого:", StudentCardNumberText.Text, labelFont, boldFont);

            document.Add(infoTable);

            Paragraph place = new Paragraph($"г. Москва, {DateTime.Now.Year} год", smallFont);
            place.Alignment = Element.ALIGN_CENTER;
            place.SpacingBefore = 80;
            document.Add(place);
        }

        private void AddPersonalDataTable(Document document, Font normalFont, Font boldFont, Font labelFont)
        {
            PdfPTable table = new PdfPTable(4);
            table.WidthPercentage = 100;
            table.SpacingBefore = 15;
            table.SpacingAfter = 20;
            table.SetWidths(new float[] { 20f, 30f, 20f, 30f });

            AddPdfRow(table, "Фамилия:", GetLastName(), "Имя:", GetFirstName(), labelFont, boldFont);
            AddPdfRow(table, "Отчество:", GetMiddleName(), "Дата рождения:", BirthDateText.Text, labelFont, normalFont);
            AddPdfRow(table, "Место рождения:", BirthPlaceText.Text, "Пол:", GenderText.Text, labelFont, normalFont);
            AddPdfRow(table, "Национальность:", NationalityText.Text, "Гражданство:", CitizenshipText.Text, labelFont, normalFont);
            AddPdfRow(table, "Телефон:", PhoneText.Text, "Email:", EmailText.Text, labelFont, normalFont);
            AddPdfRow(table, "Телефон родителей:", ParentsPhoneText.Text, "", "", labelFont, normalFont);
            AddPdfRow(table, "Адрес регистрации:", RegistrationAddressText.Text, "", "", labelFont, normalFont);
            AddPdfRow(table, "Адрес проживания:", ResidentialAddressText.Text, "", "", labelFont, normalFont);

            document.Add(table);
        }

        private void AddEducationSection(Document document, Font normalFont, Font boldFont, Font labelFont)
        {
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 100;
            table.SpacingBefore = 15;
            table.SpacingAfter = 20;
            table.SetWidths(new float[] { 30f, 70f });

            AddPdfRow(table, "Образование до поступления:", EducationBeforeText.Text, labelFont, normalFont);
            AddPdfRow(table, "Документ об образовании:", EducationDocumentText.Text, labelFont, normalFont);
            AddPdfRow(table, "Дата зачисления:", EnrollmentDateText.Text, labelFont, normalFont);
            AddPdfRow(table, "Дата окончания:", GraduationDateText.Text, labelFont, normalFont);

            document.Add(table);
        }

        private void AddSocialStatusSection(Document document, Font normalFont, Font boldFont)
        {
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 100;
            table.SpacingBefore = 15;
            table.SpacingAfter = 20;

            AddPdfRow(table, "Социальный статус:", SocialStatusText.Text, boldFont, normalFont);
            AddPdfRow(table, "Сирота:", IsOrphanCheck.IsChecked == true ? "Да" : "Нет", normalFont, normalFont);
            AddPdfRow(table, "Инвалид:", IsDisabledCheck.IsChecked == true ? "Да" : "Нет", normalFont, normalFont);
            AddPdfRow(table, "Многодетная семья:", IsFromLargeFamilyCheck.IsChecked == true ? "Да" : "Нет", normalFont, normalFont);
            AddPdfRow(table, "Малообеспеченный:", IsLowIncomeCheck.IsChecked == true ? "Да" : "Нет", normalFont, normalFont);

            document.Add(table);
        }

        private void AddEmploymentSection(Document document, Font normalFont, Font boldFont)
        {
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 100;
            table.SpacingBefore = 15;
            table.SpacingAfter = 20;

            AddPdfRow(table, "Трудоустроен:", IsEmployedCheck.IsChecked == true ? "Да" : "Нет", normalFont, normalFont);

            if (IsEmployedCheck.IsChecked == true)
            {
                AddPdfRow(table, "Место работы:", WorkPlaceText.Text, normalFont, normalFont);
                AddPdfRow(table, "Должность:", WorkPositionText.Text, normalFont, normalFont);
            }

            document.Add(table);
        }

        private void AddGradesSection(Document document, Font normalFont, Font boldFont, Font smallFont)
        {
            var grades = GradesDataGrid.ItemsSource as List<GradeModel>;
            if (grades != null && grades.Any())
            {
                PdfPTable table = new PdfPTable(4);
                table.WidthPercentage = 100;
                table.SetWidths(new float[] { 40f, 15f, 20f, 25f });

                table.AddCell(GetPdfCell("Предмет", boldFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Оценка", boldFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Семестр", boldFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Преподаватель", boldFont, Element.ALIGN_CENTER, Rectangle.BOX));

                foreach (var grade in grades)
                {
                    table.AddCell(GetPdfCell(grade.Subject ?? "-", normalFont, Element.ALIGN_LEFT, Rectangle.BOX));
                    table.AddCell(GetPdfCell(grade.Grade?.ToString() ?? "-", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                    table.AddCell(GetPdfCell(grade.Semester?.ToString() ?? "-", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                    table.AddCell(GetPdfCell(grade.Teacher ?? "-", normalFont, Element.ALIGN_LEFT, Rectangle.BOX));
                }

                document.Add(table);
            }
            else
            {
                document.Add(new Paragraph("Нет данных об успеваемости", normalFont));
            }
        }

        private void AddAchievementsSection(Document document, Font normalFont, Font smallFont)
        {
            var achievements = AchievementsDataGrid.ItemsSource as List<AchievementModel>;
            if (achievements != null && achievements.Any())
            {
                PdfPTable table = new PdfPTable(4);
                table.WidthPercentage = 100;
                table.SetWidths(new float[] { 40f, 20f, 20f, 20f });

                table.AddCell(GetPdfCell("Название", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Тип", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Дата", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Результат", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));

                foreach (var ach in achievements)
                {
                    table.AddCell(GetPdfCell(ach.AchievementName ?? "-", normalFont, Element.ALIGN_LEFT, Rectangle.BOX));
                    table.AddCell(GetPdfCell(ach.AchievementType ?? "-", normalFont, Element.ALIGN_LEFT, Rectangle.BOX));
                    table.AddCell(GetPdfCell(ach.Date?.ToString("dd.MM.yyyy") ?? "-", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                    table.AddCell(GetPdfCell(ach.Place ?? "-", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                }

                document.Add(table);
            }
            else
            {
                document.Add(new Paragraph("Нет данных о достижениях", normalFont));
            }
        }

        private void AddEventsSection(Document document, Font normalFont, Font smallFont)
        {
            var events = EventsDataGrid.ItemsSource as List<EventModel>;
            if (events != null && events.Any())
            {
                PdfPTable table = new PdfPTable(4);
                table.WidthPercentage = 100;
                table.SetWidths(new float[] { 40f, 20f, 20f, 20f });

                table.AddCell(GetPdfCell("Мероприятие", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Тип", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Дата", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                table.AddCell(GetPdfCell("Роль", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));

                foreach (var ev in events)
                {
                    table.AddCell(GetPdfCell(ev.EventName ?? "-", normalFont, Element.ALIGN_LEFT, Rectangle.BOX));
                    table.AddCell(GetPdfCell(ev.EventType ?? "-", normalFont, Element.ALIGN_LEFT, Rectangle.BOX));
                    table.AddCell(GetPdfCell(ev.Date?.ToString("dd.MM.yyyy") ?? "-", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                    table.AddCell(GetPdfCell(ev.Role ?? "-", normalFont, Element.ALIGN_CENTER, Rectangle.BOX));
                }

                document.Add(table);
            }
            else
            {
                document.Add(new Paragraph("Нет данных об участии в мероприятиях", normalFont));
            }
        }

        private void AddParentsSection(Document document, Font normalFont, Font boldFont)
        {
            var parents = ParentsDataGrid.ItemsSource as List<ParentModel>;
            if (parents != null && parents.Any())
            {
                foreach (var parent in parents)
                {
                    PdfPTable table = new PdfPTable(2);
                    table.WidthPercentage = 100;
                    table.SpacingBefore = 10;
                    table.SpacingAfter = 15;
                    table.SetWidths(new float[] { 30f, 70f });

                    AddPdfRow(table, "ФИО:", parent.FullName ?? "-", boldFont, normalFont);
                    AddPdfRow(table, "Степень родства:", parent.Relationship ?? "-", normalFont, normalFont);
                    AddPdfRow(table, "Телефон:", parent.Phone ?? "-", normalFont, normalFont);
                    AddPdfRow(table, "Место работы:", parent.WorkPlace ?? "-", normalFont, normalFont);
                    AddPdfRow(table, "Должность:", parent.WorkPosition ?? "-", normalFont, normalFont);

                    document.Add(table);
                }
            }
            else
            {
                document.Add(new Paragraph("Нет данных о родителях", normalFont));
            }
        }

        private void AddCharacteristicSection(Document document, Font normalFont)
        {
            Paragraph characteristic = new Paragraph(CharacteristicTextBox.Text, normalFont);
            characteristic.SpacingBefore = 15;
            characteristic.SpacingAfter = 20;
            document.Add(characteristic);

            Paragraph signature = new Paragraph($"\n\nКуратор: __________________\n" +
                                               $"«___» __________ {DateTime.Now.Year} г.", normalFont);
            signature.SpacingBefore = 30;
            signature.Alignment = Element.ALIGN_RIGHT;
            document.Add(signature);
        }

        private void AddFooterSection(Document document, Font smallFont, Font boldFont)
        {
            document.NewPage();

            Paragraph header = new Paragraph("ЗАКЛЮЧИТЕЛЬНАЯ ЧАСТЬ", boldFont);
            header.SpacingAfter = 20;
            document.Add(header);

            PdfPTable signatures = new PdfPTable(3);
            signatures.WidthPercentage = 100;
            signatures.SpacingBefore = 150;

            signatures.AddCell(GetPdfCell("Куратор группы", smallFont, Element.ALIGN_CENTER, Rectangle.NO_BORDER));
            signatures.AddCell(GetPdfCell("Заведующий отделением", smallFont, Element.ALIGN_CENTER, Rectangle.NO_BORDER));
            signatures.AddCell(GetPdfCell("Директор", smallFont, Element.ALIGN_CENTER, Rectangle.NO_BORDER));

            signatures.AddCell(GetPdfCell("_________________", smallFont, Element.ALIGN_CENTER, Rectangle.NO_BORDER));
            signatures.AddCell(GetPdfCell("_________________", smallFont, Element.ALIGN_CENTER, Rectangle.NO_BORDER));
            signatures.AddCell(GetPdfCell("_________________", smallFont, Element.ALIGN_CENTER, Rectangle.NO_BORDER));

            document.Add(signatures);

            Paragraph dateInfo = new Paragraph($"Документ сформирован: {DateTime.Now:dd.MM.yyyy HH:mm}", smallFont);
            dateInfo.Alignment = Element.ALIGN_CENTER;
            dateInfo.SpacingBefore = 50;
            document.Add(dateInfo);
        }

        // Вспомогательные методы для PDF
        private PdfPCell GetPdfCell(string text, Font font, int alignment, int border)
        {
            PdfPCell cell = new PdfPCell(new Phrase(text, font));
            cell.HorizontalAlignment = alignment;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Border = border;
            cell.Padding = 5;
            return cell;
        }

        private void AddPdfRow(PdfPTable table, string label, string value, Font labelFont, Font valueFont)
        {
            table.AddCell(GetPdfCell(label, labelFont, Element.ALIGN_RIGHT, Rectangle.NO_BORDER));
            table.AddCell(GetPdfCell(value, valueFont, Element.ALIGN_LEFT, Rectangle.NO_BORDER));
        }

        private void AddPdfRow(PdfPTable table, string label1, string value1, string label2, string value2,
            Font labelFont, Font valueFont)
        {
            table.AddCell(GetPdfCell(label1, labelFont, Element.ALIGN_RIGHT, Rectangle.NO_BORDER));
            table.AddCell(GetPdfCell(value1, valueFont, Element.ALIGN_LEFT, Rectangle.NO_BORDER));
            table.AddCell(GetPdfCell(label2, labelFont, Element.ALIGN_RIGHT, Rectangle.NO_BORDER));
            table.AddCell(GetPdfCell(value2, valueFont, Element.ALIGN_LEFT, Rectangle.NO_BORDER));
        }

        private void AddSectionHeader(Document document, string title, Font headerFont)
        {
            Paragraph header = new Paragraph(title, headerFont);
            header.SpacingBefore = 10;
            header.SpacingAfter = 5;
            document.Add(header);

            Paragraph line = new Paragraph(new Chunk(new LineSeparator(0.5f, 100f, BaseColor.BLACK, Element.ALIGN_CENTER, -2)));
            line.SpacingAfter = 10;
            document.Add(line);
        }

        // Методы для работы с ФИО
        private string GetLastName()
        {
            var parts = StudentFullNameText.Text.Split(' ');
            return parts.Length > 0 ? parts[0] : "";
        }

        private string GetFirstName()
        {
            var parts = StudentFullNameText.Text.Split(' ');
            return parts.Length > 1 ? parts[1] : "";
        }

        private string GetMiddleName()
        {
            var parts = StudentFullNameText.Text.Split(' ');
            return parts.Length > 2 ? parts[2] : "";
        }

        #endregion

        private void AssignHeadmanButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Проверяем, не является ли уже старостой
                if (IsHeadmanCheckBox.IsChecked == true)
                {
                    MessageBox.Show("Этот студент уже является старостой", "Информация",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Снимаем старосту с текущего в этой группе
                var student = db.Students.Find(_studentId);
                if (student == null) return;

                var currentHeadman = db.Students
                    .FirstOrDefault(s => s.GroupID == student.GroupID && s.IsHeadman == true);

                if (currentHeadman != null)
                {
                    currentHeadman.IsHeadman = false;
                }

                // Назначаем нового
                student.IsHeadman = true;
                db.SaveChanges();

                IsHeadmanCheckBox.IsChecked = true;
                MessageBox.Show($"Студент назначен старостой группы", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}