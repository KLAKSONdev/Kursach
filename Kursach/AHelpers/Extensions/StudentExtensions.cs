using Kursach.AModels;
using System;
using System.Collections.Generic;
using System.Linq;  

namespace Kursach.AHelpers.Extensions
{
    /// <summary>
    /// Extension методы для работы со студентами
    /// </summary>
    public static class StudentExtensions
    {

        public static string GetFullName(this Students student)
        {
            return $"{student.LastName} {student.FirstName} {student.MiddleName}".Trim();
        }


        public static int? CalculateAge(this Students student)
        {
            if (!student.BirthDate.HasValue) return null;

            var today = DateTime.Today;
            var birthDate = student.BirthDate.Value;
            var age = today.Year - birthDate.Year;

            if (birthDate.Date > today.AddYears(-age))
                age--;

            return age;
        }


        public static string BuildSocialStatus(this Students student)
        {
            var statuses = new List<string>();

            if (student.IsOrphan == true) statuses.Add("Сирота");
            if (student.IsDisabled == true) statuses.Add("Инвалид");
            if (student.IsFromLargeFamily == true) statuses.Add("Многодетная семья");
            if (student.IsLowIncome == true) statuses.Add("Малообеспеченный");

            return statuses.Any() ? string.Join(", ", statuses) : "Не указан";
        }


        public static bool IsHeadman(this Students student)
        {
            return student.IsHeadman ?? false;
        }


        public static bool IsActive(this Students student)
        {
            return student.IsActive ?? true;
        }
    }
}