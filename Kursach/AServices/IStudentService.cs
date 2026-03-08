using Kursach.AModels.DTO;
using System.Collections.Generic;

namespace Kursach.AServices
{
    public interface IStudentService
    {
        /// <summary>
        /// Получить всех активных студентов
        /// </summary>
        List<StudentDto> GetAllStudents();

        List<StudentDto> GetStudentsByGroup(int groupId);

        StudentDto GetStudentById(int studentId);

        bool AddStudent(StudentDto student);

        bool UpdateStudent(StudentDto student);

        bool DeleteStudent(int studentId);

        List<StudentDto> SearchStudents(string searchTerm);

        List<StudentDto> GetStudentsForRole(string role, int? groupId = null);
    }
}