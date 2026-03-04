using System;

namespace Kursach.AModels
{
    public class StudentViewModel
    {
        public int StudentID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string FullName { get; set; }
        public int? GroupID { get; set; }
        public string GroupName { get; set; }
        public int? Course { get; set; }
        public string StudentCardNumber { get; set; }
        public string PersonalNumber { get; set; }
        public DateTime? BirthDate { get; set; }
        public string BirthPlace { get; set; }
        public string Gender { get; set; }
        public string Nationality { get; set; }
        public string Citizenship { get; set; }
        public int? Age { get; set; }
        public string EducationBefore { get; set; }
        public string EducationDocument { get; set; }
        public string PhotoPath { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string ParentsPhone { get; set; }
        public string RegistrationAddress { get; set; }
        public string ResidentialAddress { get; set; }
        public int? HealthStatusID { get; set; }
        public int? FamilyTypeID { get; set; }
        public bool IsOrphan { get; set; }
        public bool IsDisabled { get; set; }
        public bool IsFromLargeFamily { get; set; }
        public bool IsLowIncome { get; set; }
        public string SocialStatus { get; set; }
        public bool IsEmployed { get; set; }
        public string WorkPlace { get; set; }
        public string WorkPosition { get; set; }
        public string Login { get; set; }
        public bool IsActive { get; set; }
        public DateTime? EnrollmentDate { get; set; }
        public DateTime? GraduationDate { get; set; }
        public DateTime? CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public bool IsHeadman { get; set; }
    }
}