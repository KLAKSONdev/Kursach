using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kursach.AModels
{
    public class EventListItem
    {
        public int EventID { get; set; }
        public string EventName { get; set; }
        public string EventType { get; set; }
        public DateTime EventDate { get; set; }
        public string Location { get; set; }
        public int ParticipantsCount { get; set; }
    }
}
