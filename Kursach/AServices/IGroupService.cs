using Kursach.AModels.DTO;
using System.Collections.Generic;

namespace Kursach.AServices
{
    public interface IGroupService
    {
        List<GroupDto> GetAllGroups();
        List<string> GetGroupNames();
        GroupDto GetGroupById(int groupId);
        int GetStudentCount(int groupId);
    }
}