using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class Station
    {
        public string stationName { get; set; }
        //0-普通-有停有发，1-始发-有被接续有发，2-终到-有到有被接续，3-通过
       public int stationType { get; set; }
        public string stoppedTime { get; set; }
        public string startedTime { get; set; }
        public string stationTrackNum { get; set; }
    }
}
