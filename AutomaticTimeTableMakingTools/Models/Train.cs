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
        //上下行 true↓ false↑
        public bool upOrDown { get; set; }
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
            stopStation = "";
            upOrDown = false;
            mainStation = new Station();
            newStations = new List<Station>();
            shownInFiles = new List<TrainFile>();
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
            if (reg.IsMatch(mainStation.startedTime))
            {//有中文，则有接续
                thisStartedTime = mainStation.stoppedTime.Replace(":", "").Trim();
            }
            else
            {
                thisStartedTime = mainStation.startedTime;
            }
            if (reg.IsMatch(otherTrain.mainStation.startedTime))
            {
                otherStartedTime = otherTrain.mainStation.stoppedTime.Replace(":", "").Trim();
            }
            else
            {
                otherStartedTime = otherTrain.mainStation.startedTime;
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
