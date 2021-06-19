using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class Station
    {
        public string stationName { get; set; }
        //0-普通-有停有发，1-始发-有被接续有发，2-终到-有到有被接续，3-通过，-1为排序用车站不打印，

       public int stationType { get; set; }
        //新增该车站在列车中与主站的关系，1为运行方向前方，0为运行方向后方
        public int stationDirection { get; set; }
        public string stoppedTime { get; set; }
        public string startedTime { get; set; }
        public string stationTrackNum { get; set; }

        public Station()
        {
            stationName = "";
            stationType = 0;
            stoppedTime = "";
            startedTime = "";
            stationTrackNum = "";
            stationDirection = -1;
        }

        public Station(Station _s)
        {
            stationName = _s.stationName;
            stationType = _s.stationType;
            stoppedTime = _s.stoppedTime;
            startedTime = _s.startedTime;
            stationTrackNum = _s.stationTrackNum;
            stationDirection = -_s.stationDirection;
        }
    }
}
