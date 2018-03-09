using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class Train
    {
        public string firstTrainNum { get; set; }
        public string secondTrainNum { get; set; }
        //始发A-B终到
        public string startStation { get; set; }
        public string stopStation { get; set; }
        //上下行
        public bool upOrDown { get; set; }
        public List<Station> newStations { get; set; }
        public List<Station> currentStations { get; set; }
        public List<FileInfo> shownInFiles { get; set; }
    }
}
