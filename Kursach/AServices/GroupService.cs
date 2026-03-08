using Kursach.AModels.DTO;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;

namespace Kursach.AServices
{
    public class GroupService : IGroupService
    {
        private readonly vsstuEntities _db;

        public GroupService()
        {
            _db = new vsstuEntities();
        }

        public List<GroupDto> GetAllGroups()
        {
            try
            {
                return _db.Groups
                    .Include(g => g.Specialties)
                    .Include(g => g.Specialties.Faculties)
                    .Select(g => new GroupDto
                    {
                        GroupID = g.GroupID,
                        GroupName = g.GroupName,
                        Course = g.Course,
                        SpecialtyName = g.Specialties.SpecialtyName,
                        FacultyName = g.Specialties.Faculties.FacultyName,
                        StudentCount = g.Students.Count(s => s.IsActive == true),
                        FormOfEducation = g.FormOfEducation,
                        AcademicYear = g.AcademicYear
                    })
                    .ToList();
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при получении списка групп", ex);
            }
        }

        public List<string> GetGroupNames()
        {
            try
            {
                return _db.Groups
                    .Where(g => g.Students.Any(s => s.IsActive == true))
                    .Select(g => g.GroupName)
                    .Distinct()
                    .OrderBy(g => g)
                    .ToList();
            }
            catch (Exception ex)
            {
                throw new ServiceException("Ошибка при получении названий групп", ex);
            }
        }

        public GroupDto GetGroupById(int groupId)
        {
            throw new NotImplementedException();
        }

        public int GetStudentCount(int groupId)
        {
            try
            {
                return _db.Students.Count(s => s.GroupID == groupId && s.IsActive == true);
            }
            catch (Exception ex)
            {
                throw new ServiceException($"Ошибка при подсчете студентов в группе {groupId}", ex);
            }
        }

        public void Dispose()
        {
            _db?.Dispose();
        }
    }
}