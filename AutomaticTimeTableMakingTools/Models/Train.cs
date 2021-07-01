using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace AutomaticTimeTableMakingTools.Models
{
    public class Train : IComparable<Train>
    {
        public string firstTrainNum { get; set; }
        public string secondTrainNum { get; set; }
        //始发A-B终到
        public string startStation { get; set; }
        public string stopStation { get; set; }
        //上下行 true↓ false↑(折角车显示折角前方向)
        public bool upOrDown { get; set; }
        public int upOrDownBeforeTurnAround { get; set; }
        //上下行都填写一份
        public bool bothUpAndDown { get; set; }
        //主站标签，徐兰场/京广场/城际场等，用于填时刻表确定位置
        public Station mainStation { get; set; }
        public List<Station> newStations { get; set; }
        public List<TrainFile> shownInFiles { get; set; }

        public Train()
        {
            firstTrainNum = "";
            secondTrainNum = "";
            startStation = "";
            stopStation = "";
            upOrDown = false;
            bothUpAndDown = false;
            upOrDownBeforeTurnAround = -1;
            mainStation = new Station();
            newStations = new List<Station>();
            shownInFiles = new List<TrainFile>();
                
        }

        public Train(string _firstTrainNum = "", string _secondTrainNum = "", string _startStation = "",
                                string _stopStation = "", bool _upOrDown = false, Station _mainStation = null, List<Station> _newStations = null, List<TrainFile> _shownInFiles = null, bool _bothUpAndDown = false,int _beforeTurnAround=-1)
        {
            firstTrainNum = _firstTrainNum;
            secondTrainNum = _secondTrainNum;
            startStation = _startStation;
            stopStation = _stopStation;
            upOrDown = _upOrDown;
            bothUpAndDown = _bothUpAndDown;
            upOrDownBeforeTurnAround = _beforeTurnAround;

            if(_mainStation == null)
            {
                mainStation = new Station();
            }
            else
            {
                mainStation = new Station(_mainStation);
            }

            if (_newStations == null)
            {
                newStations = new List<Station>();
            }
            else
            {
                List<Station> _tempStation = new List<Station>();
                foreach(Station _s in _newStations)
                {
                    _tempStation.Add(new Station(_s));
                }
                newStations = _tempStation;
            }

            if (_shownInFiles == null)
            {
                shownInFiles = new List<TrainFile>();
            }
            else
            {
                shownInFiles = _shownInFiles;
            }

    }

        public Train Clone()
        {
            Train _t = new Train(this.firstTrainNum, this.secondTrainNum, this.startStation, this.stopStation, this.upOrDown, this.mainStation, this.newStations, this.shownInFiles,this.bothUpAndDown,this.upOrDownBeforeTurnAround);
            return _t;
        }



        //重写的CompareTo方法，根据Id排序
        public int CompareTo(Train otherTrain)
        {
            /*
            if (null == otherTrain)
            {
                return 1;//空值比较大，返回1
            }
            //return this.Id.CompareTo(other.Id);//升序
            return this.mainStation.startedTime.CompareTo(otherTrain.mainStation.startedTime);//降序
            */
            //判断一下发车时间有没有汉字，有汉字说明是接续，此时使用终到时间进行排序。
            string thisStartedTime = "";
            string otherStartedTime = "";
            Regex reg = new Regex(@"[\u4e00-\u9fa5]");
            if (reg.IsMatch(mainStation.startedTime) || mainStation.startedTime.Contains("--"))
            {//有中文，则有接续
                thisStartedTime = mainStation.stoppedTime.Replace(":", "").Trim();
            }
            else
            {
                thisStartedTime = mainStation.startedTime.Replace(":", "").Trim();
            }
            if (reg.IsMatch(otherTrain.mainStation.startedTime) || otherTrain.mainStation.startedTime.Contains("--"))
            {
                otherStartedTime = otherTrain.mainStation.stoppedTime.Replace(":", "").Trim();
            }
            else
            {
                otherStartedTime = otherTrain.mainStation.startedTime.Replace(":", "").Trim();
            }

            if (mainStation == null || otherTrain.mainStation == null)
                throw new ArgumentException("Parameters can't be null");
            char[] arr1 = thisStartedTime.ToCharArray();
            char[] arr2 = otherStartedTime.ToCharArray();
            int i = 0, j = 0;
            while (i < arr1.Length && j < arr2.Length)
            {
                if (char.IsDigit(arr1[i]) && char.IsDigit(arr2[j]))
                {
                    string s1 = "", s2 = "";
                    while (i < arr1.Length && char.IsDigit(arr1[i]))
                    {
                        s1 += arr1[i];
                        i++;
                    }
                    while (j < arr2.Length && char.IsDigit(arr2[j]))
                    {
                        s2 += arr2[j];
                        j++;
                    }
                    if (int.Parse(s1) > int.Parse(s2))
                    {
                        return 1;
                    }
                    if (int.Parse(s1) < int.Parse(s2))
                    {
                        return -1;
                    }
                }
                else
                {
                    if (arr1[i] > arr2[j])
                    {
                        return 1;
                    }
                    if (arr1[i] < arr2[j])
                    {
                        return -1;
                    }
                    i++;
                    j++;
                }
            }
            if (arr1.Length == arr2.Length)
            {
                return 0;
            }
            else
            {
                return arr1.Length > arr2.Length ? 1 : -1;
            }
            //            return string.Compare( fileA, fileB );
            //            return( (new CaseInsensitiveComparer()).Compare( y, x ) );
        }
    }
}
