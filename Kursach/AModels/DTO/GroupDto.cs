namespace Kursach.AModels.DTO
{
    /// <summary>
    /// DTO для групп
    /// </summary>
    public class GroupDto
    {
        public int GroupID { get; set; }
        public string GroupName { get; set; }
        public int? Course { get; set; }
        public string SpecialtyName { get; set; }
        public string FacultyName { get; set; }
        public int StudentCount { get; set; }
        public string FormOfEducation { get; set; }
        public string AcademicYear { get; set; }
    }
}