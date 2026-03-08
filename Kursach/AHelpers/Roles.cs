using System.Linq;

namespace Kursach.AHelpers.Constants
{
    /// <summary>
    /// Роли пользователей в системе
    /// </summary>
    public static class Roles
    {
        public const string Administrator = "Администратор";
        public const string Curator = "Куратор";
        public const string Headman = "Староста";

        public static readonly string[] AllRoles = { Administrator, Curator, Headman };

        public static bool IsValidRole(string role) => AllRoles.Contains(role);

        public static bool CanEditStudents(string role) => role == Administrator || role == Curator;
        public static bool CanViewAllGroups(string role) => role == Administrator;
        public static bool IsHeadmanOnly(string role) => role == Headman;
    }
}