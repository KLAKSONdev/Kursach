namespace Kursach.AModels
{
    public class EventParticipant
    {
        public int StudentID { get; set; }
        public string FullName { get; set; }
        public string GroupName { get; set; }
        public bool IsParticipating { get; set; }
        public string Role { get; set; }
        public string Result { get; set; }
    }
}