using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class TimeTable
    {
        public string Title;
        public String[] stations;
        public List<Train> trains = new List<Train>();
    }
}
