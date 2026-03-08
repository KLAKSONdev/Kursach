using Kursach.AHelpers.Constants;
using Kursach.AHelpers.Extensions;
using Kursach.AModels.DTO;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;

namespace Kursach.AServices
{
    /// <summary>
    /// Сервис для работы со студентами
    /// </summary>
    public class StudentService : IStudentService
    {
        private readonly vsstuEntities _db;

        public StudentService()
        {
            _db = new vsstuEntities();
        }

        public List<StudentDto> GetAllStudents()
        {
            try
            {
                return _db.Students
                    .Include(s => s.Groups)
                    .Where(s => s.IsActive == true)
                    .ToList()
                    .Select(MapToDto)
                    .ToList();
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при получении списка студентов", ex);
            }
        }

        public List<StudentDto> GetStudentsByGroup(int groupId)
        {
            try
            {
                return _db.Students
                    .Include(s => s.Groups)
                    .Where(s => s.GroupID == groupId && s.IsActive == true)
                    .ToList()
                    .Select(MapToDto)
                    .ToList();
            }
            catch (Exception ex)
            {
                throw new ServiceException($"Ошибка при получении студентов группы {groupId}", ex);
            }
        }

        public StudentDto GetStudentById(int studentId)
        {
            try
            {
                var student = _db.Students
                    .Include(s => s.Groups)
                    .FirstOrDefault(s => s.StudentID == studentId);

                return student != null ? MapToDto(student) : null;
            }
            catch (Exception ex)
            {
                throw new ServiceException($"Ошибка при получении студента ID={studentId}", ex);
            }
        }

        public List<StudentDto> GetStudentsForRole(string role, int? groupId = null)
        {
            if (role == Roles.Headman && groupId.HasValue)
            {
                return GetStudentsByGroup(groupId.Value);
            }

            return GetAllStudents();
        }

        public List<StudentDto> SearchStudents(string searchTerm)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(searchTerm))
                    return GetAllStudents();

                searchTerm = searchTerm.ToLower().Trim();

                return GetAllStudents()
                    .Where(s =>
                        (s.LastName?.ToLower().Contains(searchTerm) == true) ||
                        (s.FirstName?.ToLower().Contains(searchTerm) == true) ||
                        (s.MiddleName?.ToLower().Contains(searchTerm) == true) ||
                        (s.FullName?.ToLower().Contains(searchTerm) == true))
                    .ToList();
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при поиске студентов", ex);
            }
        }

        public bool AddStudent(StudentDto student)
        {
            throw new NotImplementedException();
        }

        public bool UpdateStudent(StudentDto student)
        {

            throw new NotImplementedException();
        }

        public bool DeleteStudent(int studentId)
        {
            try
            {
                var student = _db.Students.Find(studentId);
                if (student == null)
                    return false;

                student.IsActive = false;
                _db.SaveChanges();
                return true;
            }
            catch (Exception ex)
            {
                throw new ServiceException($"Ошибка при удалении студента ID={studentId}", ex);
            }
        }

        #region Маппинг

        private StudentDto MapToDto(Students s)
        {
            return new StudentDto
            {
                StudentID = s.StudentID,
                LastName = s.LastName ?? "",
                FirstName = s.FirstName ?? "",
                MiddleName = s.MiddleName ?? "",
                FullName = s.GetFullName(),

                GroupID = s.GroupID,
                GroupName = s.Groups?.GroupName ?? "Без группы",
                Course = s.Groups?.Course,

                StudentCardNumber = s.StudentCardNumber ?? "",
                PersonalNumber = s.PersonalNumber ?? "",

                BirthDate = s.BirthDate,
                Age = s.CalculateAge(),
                Gender = s.Gender ?? "",
                Phone = s.Phone ?? "",
                Email = s.Email ?? "",
                ParentsPhone = s.ParentsPhone ?? "",

                RegistrationAddress = s.RegistrationAddress ?? "",
                ResidentialAddress = s.ResidentialAddress ?? "",

                IsOrphan = s.IsOrphan ?? false,
                IsDisabled = s.IsDisabled ?? false,
                IsFromLargeFamily = s.IsFromLargeFamily ?? false,
                IsLowIncome = s.IsLowIncome ?? false,
                SocialStatus = s.BuildSocialStatus(),

                IsEmployed = s.IsEmployed ?? false,
                WorkPlace = s.WorkPlace ?? "",
                WorkPosition = s.WorkPosition ?? "",

                IsHeadman = s.IsHeadman(),
                IsActive = s.IsActive(),
                EnrollmentDate = s.EnrollmentDate,
                GraduationDate = s.GraduationDate
            };
        }

        #endregion

        public void Dispose()
        {
            _db?.Dispose();
        }
    }

    /// <summary>
    /// Специализированное исключение для сервисов
    /// </summary>
    public class ServiceException : Exception
    {
        public ServiceException(string message, Exception innerException)
            : base(message, innerException) { }

        public ServiceException(string message) : base(message) { }
    }
}