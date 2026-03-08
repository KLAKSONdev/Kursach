using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;
using Kursach.AModels.DTO;

namespace Kursach.AServices
{
    /// <summary>
    /// Сервис для генерации документов (Word)
    /// </summary>
    public interface IDocumentService
    {
        /// <summary>
        /// Создание справки об обучении
        /// </summary>
        void CreateStudyCertificate(StudentDto student);

        /// <summary>
        /// Создание характеристики в военкомат
        /// </summary>
        void CreateMilitaryCharacteristic(StudentDto student);

        /// <summary>
        /// Создание заявления на академический отпуск
        /// </summary>
        void CreateAcademicLeaveApplication(StudentDto student);
    }

    public class DocumentService : IDocumentService
    {
        private readonly vsstuEntities _db;

        public DocumentService()
        {
            _db = new vsstuEntities();
        }

        // =====================================================================
        // СПРАВКА ОБ ОБУЧЕНИИ
        // =====================================================================

        public void CreateStudyCertificate(StudentDto student)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word документы (*.docx)|*.docx",
                    FileName = $"Справка_об_обучении_{student.LastName}_{DateTime.Now:yyyyMMdd}.docx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    CreateStudyCertificateDocument(saveFileDialog.FileName, student);
                }
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при создании справки", ex);
            }
        }

        private void CreateStudyCertificateDocument(string filePath, StudentDto student)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false };
                doc = wordApp.Documents.Add();
                var selection = wordApp.Selection;

                // Настройка шрифта
                selection.Font.Name = "Times New Roman";
                selection.Font.Size = 14;

                // Заголовок
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Bold = 1;
                selection.Font.Size = 16;
                selection.TypeText("СПРАВКА ОБ ОБУЧЕНИИ\n");
                selection.Font.Bold = 0;
                selection.Font.Size = 14;
                selection.TypeText("\n\n");

                // Основной текст
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                selection.TypeText($"Дана {student.FullName},\n");
                selection.TypeText($"студенту группы {student.GroupName},\n");
                selection.TypeText($"в том, что он(а) действительно обучается в\n");
                selection.TypeText("ГБПОУ \"Колледж\" по специальности\n");
                selection.TypeText($"{GetStudentSpecialty(student.StudentID)}.\n\n");
                selection.TypeText($"Курс: {student.Course}\n");
                selection.TypeText($"Форма обучения: очная\n\n");
                selection.TypeText("Справка выдана для предъявления по месту требования.\n\n\n\n");

                // Подпись
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                selection.TypeText("Директор колледжа _____________ И.И. Иванов\n");
                selection.TypeText("М.П.\n\n");
                selection.TypeText($"«___» __________ {DateTime.Now.Year} г.");

                doc.SaveAs2(filePath);
                MessageBox.Show("Справка об обучении создана!", "Успех",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CleanupWordObjects(doc, wordApp);
            }
        }

        // =====================================================================
        // ХАРАКТЕРИСТИКА В ВОЕНКОМАТ
        // =====================================================================

        public void CreateMilitaryCharacteristic(StudentDto student)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word документы (*.docx)|*.docx",
                    FileName = $"Характеристика_в_военкомат_{student.LastName}_{DateTime.Now:yyyyMMdd}.docx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    CreateMilitaryCharacteristicDocument(saveFileDialog.FileName, student);
                }
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при создании характеристики", ex);
            }
        }

        private void CreateMilitaryCharacteristicDocument(string filePath, StudentDto student)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false };
                doc = wordApp.Documents.Add();
                var selection = wordApp.Selection;

                selection.Font.Name = "Times New Roman";
                selection.Font.Size = 14;

                // Заголовок
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Bold = 1;
                selection.Font.Size = 16;
                selection.TypeText("ХАРАКТЕРИСТИКА\n\n");
                selection.Font.Bold = 0;
                selection.Font.Size = 14;

                selection.TypeText($"на студента {student.FullName}\n");
                selection.TypeText($"{student.BirthDate:dd.MM.yyyy} года рождения\n\n");

                // 1. Успеваемость
                selection.Font.Bold = 1;
                selection.TypeText("1. УСПЕВАЕМОСТЬ\n");
                selection.Font.Bold = 0;
                selection.TypeText($"За время обучения проявил себя как {GetAcademicPerformanceText(student.StudentID)} студент. ");
                selection.TypeText($"Средний балл успеваемости: {GetAverageGrade(student.StudentID)}.\n\n");

                // 2. Дисциплина
                selection.Font.Bold = 1;
                selection.TypeText("2. ДИСЦИПЛИНА\n");
                selection.Font.Bold = 0;
                selection.TypeText($"Правила внутреннего распорядка {GetDisciplineText(student.StudentID)}. ");
                selection.TypeText($"Замечаний от преподавателей {GetRemarksText(student.StudentID)}.\n\n");

                // 3. Общественная активность
                selection.Font.Bold = 1;
                selection.TypeText("3. ОБЩЕСТВЕННАЯ АКТИВНОСТЬ\n");
                selection.Font.Bold = 0;
                selection.TypeText($"В общественной жизни группы и колледжа {GetActivityText(student.StudentID)}. ");
                selection.TypeText($"Количество мероприятий с участием: {GetEventsCount(student.StudentID)}.\n\n");

                // 4. Личные качества
                selection.Font.Bold = 1;
                selection.TypeText("4. ЛИЧНЫЕ КАЧЕСТВА\n");
                selection.Font.Bold = 0;
                selection.TypeText(GetPersonalQualities(student.StudentID) + "\n\n");

                // Подписи
                selection.TypeText("Куратор группы __________________ /Иванова М.И./\n\n");
                selection.TypeText("Директор колледжа __________________ /Иванов И.И./\n");
                selection.TypeText("М.П.\n\n");
                selection.TypeText($"«___» __________ {DateTime.Now.Year} г.\n");

                doc.SaveAs2(filePath);
                MessageBox.Show("Характеристика в военкомат создана!", "Успех",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CleanupWordObjects(doc, wordApp);
            }
        }

        // =====================================================================
        // ЗАЯВЛЕНИЕ НА АКАДЕМИЧЕСКИЙ ОТПУСК
        // =====================================================================

        public void CreateAcademicLeaveApplication(StudentDto student)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word документы (*.docx)|*.docx",
                    FileName = $"Заявление_академ_{student.LastName}_{DateTime.Now:yyyyMMdd}.docx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    CreateAcademicLeaveDocument(saveFileDialog.FileName, student);
                }
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при создании заявления", ex);
            }
        }

        private void CreateAcademicLeaveDocument(string filePath, StudentDto student)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false };
                doc = wordApp.Documents.Add();
                var selection = wordApp.Selection;

                selection.Font.Name = "Times New Roman";
                selection.Font.Size = 14;

                // Шапка (справа)
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                selection.TypeText("Директору колледжа\n");
                selection.TypeText("Иванову И.И.\n");
                selection.TypeText($"от студента группы {student.GroupName}\n");
                selection.TypeText($"{student.FullName}\n");
                selection.TypeText("\n");

                // Заголовок (по центру)
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.Font.Bold = 1;
                selection.Font.Size = 16;
                selection.TypeText("ЗАЯВЛЕНИЕ\n");
                selection.Font.Bold = 0;
                selection.Font.Size = 14;
                selection.TypeText("\n");

                // Основной текст
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                selection.TypeText("Прошу предоставить мне академический отпуск по семейным обстоятельствам.\n\n");
                selection.TypeText("Срок отпуска: с «___» __________ 20__ г. по «___» __________ 20__ г.\n\n");
                selection.TypeText("Обязуюсь предоставить подтверждающие документы.\n\n");
                selection.TypeText("«___» __________ 20___ г.\n\n");
                selection.TypeText("__________________\n");
                selection.TypeText("     (подпись)\n");

                doc.SaveAs2(filePath);
                MessageBox.Show("Заявление успешно создано!", "Успех",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CleanupWordObjects(doc, wordApp);
            }
        }

        // =====================================================================
        // ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ДЛЯ ХАРАКТЕРИСТИКИ
        // =====================================================================

        private string GetStudentSpecialty(int studentId)
        {
            var student = _db.Students
                .FirstOrDefault(s => s.StudentID == studentId);
            return student?.Groups?.Specialties?.SpecialtyName ?? "Информационные технологии";
        }

        private double GetAverageGrade(int studentId)
        {
            var grades = _db.AcademicPerformance
                .Where(g => g.StudentID == studentId && g.Grade.HasValue)
                .Select(g => g.Grade.Value);
            return grades.Any() ? Math.Round(grades.Average(), 2) : 4.0;
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
            int violations = _db.DisciplinaryRecords
                .Count(d => d.StudentID == studentId);
            return violations == 0 ? "соблюдает" : "иногда нарушает";
        }

        private string GetRemarksText(int studentId)
        {
            int remarks = _db.DisciplinaryRecords
                .Count(d => d.StudentID == studentId && d.RecordType == "Замечание");
            return remarks == 0 ? "не имеет" : $"имеет ({remarks})";
        }

        private string GetActivityText(int studentId)
        {
            int events = _db.EventParticipation
                .Count(ep => ep.StudentID == studentId);

            if (events > 10) return "принимает активное участие";
            if (events > 5) return "участвует";
            if (events > 0) return "эпизодически участвует";
            return "не участвует";
        }

        private int GetEventsCount(int studentId)
        {
            return _db.EventParticipation.Count(ep => ep.StudentID == studentId);
        }

        private string GetPersonalQualities(int studentId)
        {
            var traits = _db.StudentTraits
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

        // =====================================================================
        // ОСВОБОЖДЕНИЕ РЕСУРСОВ COM
        // =====================================================================

        private void CleanupWordObjects(Word.Document doc, Word.Application app)
        {
            try
            {
                if (doc != null)
                {
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                }

                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка при освобождении Word: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _db?.Dispose();
        }
    }
}