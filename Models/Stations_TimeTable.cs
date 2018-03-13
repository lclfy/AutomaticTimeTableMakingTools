using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class Stations_TimeTable
    {
        //用于标记时刻表表头中的车站数据
        public string stationName { get; set; }
        //×↑√↓
        public bool upOrDown { get; set; }
        //该车站所在列
        public int stationColumn { get; set; }
        //该车站的到达-股道-发出 所在列
        public int stoppedTimeColumn { get; set; }
        public int trackNumColumn { get; set; }
        public int startedTimeColumn { get; set; }

    }
}
