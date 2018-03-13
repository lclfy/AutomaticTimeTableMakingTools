using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class TimeTable
    {
        public string Title;
        public String[] stations;
        //上下行分开
        public List<Train> upTrains = new List<Train>();
        public List<Train> downTrains = new List<Train>();
    }
}
