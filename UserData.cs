using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserSchedulerService
{
    public class UserData
    {
        public Guid UserId { get; set; }
        public string FullName { get; set; }
        public int Segment { get; set; }
        public int LOB { get; set; }
        public int Role { get; set; }
    }

}
