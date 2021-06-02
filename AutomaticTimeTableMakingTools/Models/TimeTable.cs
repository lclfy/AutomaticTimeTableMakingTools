using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class TimeTable
    {
        public string Title { get; set; }
        //此站名用于区分列车属于哪张时刻表
        public String[] stations { get; set; }
        //时刻表中判断上下行-名称的行位置
        public int titleRow { get; set; }
        //时刻表中判断车站所在行的行位置
        public int stationRow { get; set; }
        //将现有时刻表中的车站提取出来，并且记录列位置
        public List<Stations_TimeTable> currentStations = new List<Stations_TimeTable>();
        //后期寻找时刻表文件，在此处保存文件名
        public string fileName { get; set; }
        //与workbook对应的位置
        public int timeTablePlace { get; set; }
        //上下行分开
        public List<Train> upTrains = new List<Train>();
        public List<Train> downTrains = new List<Train>();

        public TimeTable()
        {
            Title = "";
            titleRow = 0;
            stationRow = 0;
            currentStations = new List<Stations_TimeTable>();
            fileName = "";
            timeTablePlace = 0;
            upTrains = new List<Train>();
            downTrains = new List<Train>();
        }
    }
}
