using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    class upOrDown_Stations
    {
        public string stationName { get; set; }
        public List<string> upStations { get; set; }
        public List<string> downStations { get; set; }

        public upOrDown_Stations()
        {
            stationName = "";
            upStations = new List<string>();
            downStations = new List<string>();
        }
    }
}
