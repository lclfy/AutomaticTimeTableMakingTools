using AutomaticTimeTableMakingTools.Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CCWin;

namespace AutomaticTimeTableMakingTools
{
    public partial class Main : Skin_Mac
    {
        List<IWorkbook> NewTimeTablesWorkbooks;
        List<IWorkbook> CurrentTimeTablesWorkbooks;
        String[] currentTimeTableFileNames;
        List<Train> allTrains_New = new List<Train>();
        //List<TimeTable> allTimeTables = new List<TimeTable>();

        //分表文件
        List<IWorkbook> DistributedTimeTableWorkbooks;
        String[] DistributedTimeTableFileNames;
        List<TimeTable> allDistributedTimeTables = new List<TimeTable>();
        bool hasDistributedTimeTable = false;
        //新时刻表的模式，0为普通不分上下行的（仍然需要把车单独复制到同一列），1为子东临客表，2为四大表，3为原来的分上下行的
        int selectNewTimeTableMode = 0;
        
        public Main()
        {
            InitializeComponent();
            initUI();
            this.Text = "时刻表分发工具";
        }

        private void initUI()
        {
            modeSelect_cb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            modeSelect_cb.SelectedIndex = 0;
        }

        private bool ImportFiles(int fileType)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();   //显示选择文件对话框 
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "Excel 文件 |*.xlsx;*.xls";
            //openFileDialog1.InitialDirectory = Application.StartupPath + "\\时刻表\\";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            List<IWorkbook> workBooks = new List<IWorkbook>();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                int fileCount = 0;
                String fileNames = "已选择：";
                foreach(string fileName in openFileDialog1.FileNames)
                {
                    fileCount++;
                    IWorkbook workbook = null;  //新建IWorkbook对象  
                    try
                    {
                        FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                        if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
                        {
                            try
                            {
                                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
                                workBooks.Add(workbook);
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show("读取时刻表时出现错误，请重新复制时刻表或重试\n"+fileName+"\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return false;
                            }

                        }
                        else if (fileName.IndexOf(".xls") > 0) // 2003版本  
                        {
                            try
                            {
                                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                                workBooks.Add(workbook);
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show("读取时刻表时出现错误，请重新复制时刻表或重试\n" + fileName + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return false;
                            }
                        }
                        fileStream.Close();
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("选中的部分时刻表文件正在使用中，请关闭后重试\n" + fileName, "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return false;
                    }
                }
                fileNames = fileNames + fileCount.ToString() + "个文件";
                switch (fileType)
                {//0为新，1为当前
                    case 0:
                        NewTimeTableFile_lbl.Text = fileNames;
                        NewTimeTablesWorkbooks = workBooks;
                        break;
                    case 2:
                        DistributedTimeTableFile_lbl.Text = fileNames;
                        DistributedTimeTableWorkbooks = workBooks;
                        List<string> strList_DT = new List<string>();
                        foreach (string fileName in openFileDialog1.FileNames)
                        {
                            strList_DT.Add(fileName);//循环添加元素
                        }
                        DistributedTimeTableFileNames = strList_DT.ToArray();
                        break;
                }
            }
            
            return true;
        }

        private List<TimeTable>  GetStationsFromCurrentTables(List<IWorkbook> _timeTablesWorkbooks,List<TimeTable> _allTimeTables,int _inputType)
        {
            //通过标题寻找车站（线路所模糊匹配，东三场南站动车所精确匹配）
            List<TimeTable> _timeTables = new List<TimeTable>();
            int counter = 0;
            foreach (IWorkbook workbook in _timeTablesWorkbooks)
            {
                TimeTable _timeTable = new TimeTable();
                string allStations = "";
                ISheet sheet = workbook.GetSheetAt(0);  //获取工作表  
                IRow row;

                for (int i = 0; i < sheet.LastRowNum; i++)  //对工作表每一行  
                {
                    row = sheet.GetRow(i);   //row读入第i行数据  
                    bool hasGotStationsRow = false;
                    if (row != null)
                    {
                        for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列  
                        {
                            if (row.GetCell(j) != null)
                            {
                                string titleName = "";
                                bool hasgot = false;
                                titleName = row.GetCell(j).ToString().Trim().Replace(" ","");
                                if(_inputType == 0)
                                {
                                    if ((titleName.Contains("郑州东站") &&
                                   titleName.Contains("上行")) ||
                                  (titleName.Contains("郑州东站") &&
                                    titleName.Contains("下行")))
                                    {
                                        _timeTable.Title = "郑州东";
                                        _timeTable.titleRow = i;
                                        hasgot = true;
                                    }
                                }
                                else if(_inputType == 1)
                                {
                                    string[] _allStations = new string[] { "曹古寺", "二郎庙", "鸿宝", "郑州东京广场", "南曹", "寺后", "郑州东徐兰场", "郑开", "郑州南郑万场", "郑州东城际场", "郑州南城际场", "郑州东疏解区", "郑州东动车所", "郑州南动车所" };
                                    for (int ij = 0; ij < _allStations.Length; ij++)
                                    {
                                        if ((titleName.Contains(_allStations[ij]) &&
                                                                 titleName.Contains("上行")) ||
                                                                (titleName.Contains(_allStations[ij]) &&
                                                                titleName.Contains("下行")))
                                        {
                                            //表头与车站不符的特殊情况
                                            if (titleName.Contains("郑开"))
                                            {
                                                _timeTable.Title = "宋城路";
                                            }
                                            else
                                            {
                                                _timeTable.Title = _allStations[ij];
                                            }
                                            _timeTable.titleRow = i;
                                            break;
                                        }
                                    }
                                }
                                if (row.GetCell(j).ToString().Contains("始发") ||
                                    row.GetCell(j).ToString().Contains("备注") ||
                                     hasGotStationsRow)
                                {
                                    hasGotStationsRow = true;
                                    _timeTable.stationRow = i;
                                }
                                if (!row.GetCell(j).ToString().Trim().Replace(" ","").Contains("时刻")&&
                                    row.GetCell(j).ToString().Length != 0)
                                {
                                    string currentStation = row.GetCell(j).ToString();
                                        if (currentStation.Contains("线路所"))
                                    {
                                        currentStation = currentStation.Replace("线路所", "");
                                    }
                                    if (currentStation.Contains("车站"))
                                    {
                                        currentStation = currentStation.Replace("车站", "车次");
                                    }
                                        if (currentStation.Contains("站"))
                                    {
                                        currentStation = currentStation.Replace("站", "");
                                    }

                                        if (currentStation.Contains("郑州东"))
                                    {
                                        //郑州南城际场修改
                                        /*
                                        if (currentStation.Equals("郑州东城际场"))
                                        {
                                            continue;
                                        }
                                        */
                                        if (_inputType == 0)
                                        {
                                            currentStation = "郑州东";
                                        }
                                        else if(_inputType == 1)
                                        {
                                            if (currentStation.Contains("郑州东京广场"))
                                            {
                                                currentStation = "郑州东京广场";
                                            }
                                            if (currentStation.Contains("郑州东城际场"))
                                            {
                                                currentStation = "郑州东城际场";
                                            }
                                            if (currentStation.Contains("郑州东徐兰场"))
                                            {
                                                currentStation = "郑州东徐兰场";
                                            }
                                        }
                                        //currentStation = currentStation.Replace("郑州东", "");
                                    } 
                                    currentStation = currentStation.Trim();
                                    Stations_TimeTable _tempStation = new Stations_TimeTable();
                                    //此时需要找这趟车是上行还是下行
                                    IRow titleRow = sheet.GetRow(_timeTable.titleRow);
                                    _tempStation.stationColumn = j;
                                    _tempStation.stationName = currentStation;
                                    if (titleRow != null)
                                    {
                                        bool hasGotData = false;
                                        for (int k = j; k >= 0; k--)
                                        {//往上找写了上下行的那行，往左找 直到找到字为止
                                            if (titleRow.GetCell(k) != null)
                                            {
                                                string cellInfo = titleRow.GetCell(k).ToString();
                                                if (cellInfo.Contains("上行"))
                                                {//说明是上行的
                                                    _tempStation.upOrDown = false;
                                                    hasGotData = true;
                                                }
                                                else if (cellInfo.Contains("下行"))
                                                {
                                                    _tempStation.upOrDown = true;
                                                    hasGotData = true;
                                                }
                                            }
                                            if (hasGotData)
                                            {
                                                break;
                                            }
                                        }
                                        if (!allStations.Contains(currentStation))
                                        {
                                            allStations = allStations + "-"+ currentStation;
                                        }
                                        
                                    }
                                    else
                                    {
                                        MessageBox.Show("选定的列车时刻表表头不具有规定格式：“郑州东站…时刻表（上行）”或“（线路所）…时刻表（上行）”", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        return null;
                                    }
                                    //此时依然不能直接添加，需要寻找到达-股道-发出所在列
                                    IRow stopStartRow = sheet.GetRow(_timeTable.stationRow + 1);
                                    if(stopStartRow != null)
                                    {
                                        string cellInfo = "";
                                        if(stopStartRow.GetCell(j) != null)
                                        {
                                            cellInfo = stopStartRow.GetCell(j).ToString().Trim();

                                                if (cellInfo.Contains("通过") || cellInfo.Contains("发出"))
                                                {
                                                    _tempStation.startedTimeColumn = j;
                                                }
                                                if (cellInfo.Contains("到达"))
                                                {
                                                    _tempStation.stoppedTimeColumn = j;
                                                }
                                            if (cellInfo.Contains("股道"))
                                            {
                                                _tempStation.trackNumColumn = j;
                                            }
                                                //此时往右，再往上，看看get到的是不是自己，是的话就看是股道还是发出，直到不是的再退出循环
                                                for(int k = j + 1; k < stopStartRow.LastCellNum; k++)
                                                {//
                                                    string nextCell = "";
                                                    if(stopStartRow.GetCell(k) != null)
                                                    {
                                                        nextCell = stopStartRow.GetCell(k).ToString().Trim();
                                                    }
                                                    IRow stationRow = sheet.GetRow(_timeTable.stationRow);
                                                    if(stationRow != null)
                                                    {
                                                        if(stationRow.GetCell(k) == null)
                                                        {
                                                            if (nextCell.Contains("股道"))
                                                            {
                                                                _tempStation.trackNumColumn = k;
                                                            }else if (nextCell.Contains("发出"))
                                                            {
                                                                _tempStation.startedTimeColumn = k;
                                                            }
                                                        }
                                                        else if(stationRow.GetCell(k).ToString().Length == 0)
                                                        {
                                                            if (nextCell.Contains("股道"))
                                                            {
                                                                _tempStation.trackNumColumn = k;
                                                            }
                                                            else if (nextCell.Contains("发出"))
                                                            {
                                                                _tempStation.startedTimeColumn = k;
                                                            }
                                                        }
                                                        else
                                                        {//有字就不对了，应该跳出
                                                            break;
                                                        }
                                                    }
                                                }

                                        }
                                        else
                                        {
                                            MessageBox.Show("选定的列车时刻表表头不具有规定格式：到达-股道-发出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        }

                                    }
                                    else
                                    {
                                        MessageBox.Show("选定的列车时刻表表头不具有规定格式：到达-股道-发出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        return null;
                                    }
                                    _timeTable.currentStations.Add(_tempStation);
                                }
                            }
                        }
                    }
                    if (hasGotStationsRow)
                    {
                        break;
                    }
                }
                //仅使用郑万时刻表，作为徐兰场使用
                if (_timeTable.Title.Contains("郑万") && _timeTables.Count == 1)
                {
                    _timeTable.Title = "徐兰";
                }
                else if (_timeTable.Title == null || _timeTable.Title.Length == 0)
                {
                    MessageBox.Show("选定的列车时刻表表头不具有规定格式：“郑州东站…时刻表（上行）”或“（线路所）…时刻表（上行）”", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return null;
                }
                allStations = allStations.Remove(0,1);
                _timeTable.stations = allStations.Split('-');
                if(_inputType == 0)
                {
                    _timeTable.fileName = currentTimeTableFileNames[counter];
                }
                else if(_inputType == 1)
                {
                    _timeTable.fileName = DistributedTimeTableFileNames[counter];
                }

                _timeTable.timeTablePlace = counter;
                _timeTables.Add(_timeTable);
                _allTimeTables = _timeTables;
                counter++;
            }
            //passingStations = allStations;
            string outPut = "";
            foreach(TimeTable table in _allTimeTables)
            {
                for(int i = 0; i < table.stations.Length; i++)
                {
                    outPut = outPut + table.Title + "-" + table.stations[i].ToString() + "||";
                } 
            }
            return _allTimeTables;

        }

        //子东版时刻表识别
        private List<Train> ZiDongVersion_GetTrainsFromNewTimeTables()
        {
            List<Train> allTrains = new List<Train>();
            foreach (IWorkbook workbook in NewTimeTablesWorkbooks)
            {
                ISheet sheet = null;
                string a = "";
                for (int i = 0; i< workbook.NumberOfSheets; i++)
                {
                    a = workbook.GetSheetAt(i).SheetName;
                    if (workbook.GetSheetAt(i).SheetName.Contains("汇总表"))
                    {
                        sheet = workbook.GetSheetAt(i);
                        break;
                    }
                }
                if(sheet == null)
                {
                    if(workbook.NumberOfSheets > 0)
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                    else
                    {
                        MessageBox.Show("无法找到正确的时刻表。请将临客表内包含需要导出列车的表格内容单独复制至新Excel文件内，并重新选择文件","提示",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        return allTrains;
                    }
                }
                IRow row;
                //包含“车次”的列有车次，始发站，终到站，行别，车站的到达股道发车
                //“车次”的上面一行有车站名称。
                int trainNumberColumn = -1;
                int startStationColumn = -1;
                int stopStationColumn = -1;
                int upOrDownColumn = -1;
                //标题行
                int titleRowNum = -1;
                //车站行
                int stationRowNum = -1;
                //先找标题
                for(int  _rowNum = 0; _rowNum <= sheet.LastRowNum;_rowNum++)
                {
                    if(sheet.GetRow(_rowNum) == null)
                    {
                        continue;
                    }
                    else
                    {
                        row = sheet.GetRow(_rowNum);
                    }
                    for(int  _columnNum = 0; _columnNum <= row.LastCellNum; _columnNum++)
                    {
                        if(row.GetCell(_columnNum)!= null)
                        {
                            //车次列
                            if (row.GetCell(_columnNum).ToString().Contains("车次"))
                            {
                                trainNumberColumn = _columnNum;
                                titleRowNum = _rowNum;
                                stationRowNum = _rowNum - 1;
                            }
                            if (row.GetCell(_columnNum).ToString().Contains("始发"))
                            {
                                startStationColumn = _columnNum;
                            }
                            if (row.GetCell(_columnNum).ToString().Contains("终到"))
                            {
                                stopStationColumn = _columnNum;
                            }
                            if (row.GetCell(_columnNum).ToString().Contains("行别"))
                            {
                                upOrDownColumn = _columnNum;
                            }
                        }
                    }
                }
                //开始找车
                if (trainNumberColumn != -1 &&
                    startStationColumn != -1 &&
                    stopStationColumn != -1 &&
                    upOrDownColumn != -1&&
                    titleRowNum != -1)
                {
                    //从标题的下一行开始找
                    //找到一个内容后，判断标题行是否为车次，始发站，终到站，行别，不是的话检查标题行
                    //如果是“到达”，则向上找站名，没有的话向右一个，看看是不是“股道”，再向上找，没有的话再往右一个
                    //以此类推，“股道”的话，先看正上方，再看左上方，最后看右上方
                    //“发车”的话，先看正上方，再看左上方，再看左上方
                    IRow TitleRow = sheet.GetRow(titleRowNum);
                    IRow StationRow = sheet.GetRow(stationRowNum);
                    for(int trainRowNum = titleRowNum + 1; trainRowNum <= sheet.LastRowNum; trainRowNum++)
                    {
                        //一行一趟车
                        IRow trainRow = sheet.GetRow(trainRowNum);
                        Train _train = new Train();
                        for (int trainColumn = 0;trainColumn <= trainRow.LastCellNum; trainColumn++)
                        {
                            ICell trainCell = trainRow.GetCell(trainColumn);
                            //去掉空格
                            if (trainCell != null && trainCell.ToString().Trim().Length != 0)
                            {
                                //基本数据
                                if (TitleRow.GetCell(trainColumn).ToString().Trim() == null)
                                { 
                                    continue;
                                }
                                if (trainColumn == trainNumberColumn)
                                {
                                    if (!trainCell.ToString().Trim().Contains("G") &&
                                    !trainCell.ToString().Trim().Contains("D") &&
                                    !trainCell.ToString().Trim().Contains("C") &&
                                    !trainCell.ToString().Trim().Contains("J"))
                                    {
                                        continue;
                                    }
                                    if (trainCell.ToString().Trim().Contains("/"))
                                    {
                                        string[] TrainNums = splitTrainNum(trainCell.ToString().Trim());
                                        _train.firstTrainNum = TrainNums[0];
                                        _train.secondTrainNum = TrainNums[1];
                                    }
                                    else
                                    {
                                        _train.firstTrainNum = trainCell.ToString().Trim();
                                    }

                                    continue;
                                }
                                else if(trainColumn == startStationColumn)
                                {
                                    _train.startStation = trainCell.ToString().Trim();
                                    continue;
                                }
                                else if (trainColumn == stopStationColumn)
                                {
                                    _train.stopStation = trainCell.ToString().Trim();
                                    continue;
                                }
                                else if (trainColumn == upOrDownColumn)
                                {
                                    if (trainCell.ToString().Trim().Contains("下行"))
                                    {
                                        _train.upOrDown = true;
                                    }
                                    else if (trainCell.ToString().Trim().Contains("上行"))
                                    {
                                        _train.upOrDown = false;
                                    }
                                    //新增一个参数，标记上下行不明的车
                                    if(trainCell.ToString().Trim().Contains("下上") || trainCell.ToString().Trim().Contains("上下")||
                                        trainCell.ToString().Trim().Contains("下/上") || trainCell.ToString().Trim().Contains("上/下"))
                                    {
                                        _train.hasNoUpOrDown = true;
                                    }
                                    continue;
                                }
                                //找车
                                Station _s = new Station();
                                if (TitleRow.GetCell(trainColumn).ToString().Trim().Equals("到达"))
                                {//直接把这个车这个站添加进去
                                    //先把到达时间填上
                                    _s.stoppedTime = trainRow.GetCell(trainColumn).ToString().Trim();
                                    //先找正上方有没有站名，没有的话向右找
                                    if(StationRow.GetCell(trainColumn) != null &&
                                        StationRow.GetCell(trainColumn).ToString().Trim().Length != 0)
                                    {
                                        _s.stationName = StationRow.GetCell(trainColumn).ToString().Trim();
                                    }
                                    //+1是“股道”
                                    if (TitleRow.GetCell(trainColumn + 1) != null)
                                    {
                                        if (TitleRow.GetCell(trainColumn + 1).ToString().Trim().Equals("股道"))
                                        {
                                            //如果站名还没填上，则填上
                                            if (_s.stationName.Length == 0 && StationRow.GetCell(trainColumn + 1) != null &&
                                            StationRow.GetCell(trainColumn + 1).ToString().Trim().Length != 0)
                                            {
                                                _s.stationName = StationRow.GetCell(trainColumn + 1).ToString().Trim();
                                            }
                                            //填上股道
                                            if(trainRow.GetCell(trainColumn+1) != null && trainRow.GetCell(trainColumn+1).ToString().Trim().Length != 0)
                                            {
                                                _s.stationTrackNum = trainRow.GetCell(trainColumn+1).ToString().Trim();
                                            }
                                        }
                                    }
                                    //+2是“发车”
                                    if (TitleRow.GetCell(trainColumn + 2) != null)
                                    {
                                        if (TitleRow.GetCell(trainColumn + 2).ToString().Trim().Equals("发车"))
                                        {
                                            //如果站名还没填上，则填上
                                            if (_s.stationName.Length == 0 && StationRow.GetCell(trainColumn + 2) != null &&
                                            StationRow.GetCell(trainColumn + 2).ToString().Trim().Length != 0)
                                            {
                                                _s.stationName = StationRow.GetCell(trainColumn + 2).ToString().Trim();
                                            }
                                            //填上发车点
                                            if (trainRow.GetCell(trainColumn+2) != null && trainRow.GetCell(trainColumn+2).ToString().Trim().Length != 0)
                                            {
                                                _s.startedTime = trainRow.GetCell(trainColumn+2).ToString().Trim();
                                            }
                                        }
                                    }
                                }
                                else if(TitleRow.GetCell(trainColumn).ToString().Trim().Equals("股道"))
                                {//如果左边是到达，且左边有内容，则添加过车站了，跳过
                                    if (trainRow.GetCell(trainColumn - 1) != null &&
                                        trainRow.GetCell(trainColumn - 1).ToString().Trim().Length != 0 &&
                                        TitleRow.GetCell(trainColumn  - 1) != null &&
                                    TitleRow.GetCell(trainColumn  -1).ToString().Trim().Length != 0)
                                    {
                                        if (TitleRow.GetCell(trainColumn -1).ToString().Trim().Equals("到达"))
                                        {
                                            continue;
                                        }
                                    }
                                    //没跳过，就添加 如法炮制
                                    //先把股道填上
                                    _s.stationTrackNum = trainRow.GetCell(trainColumn).ToString().Trim();
                                    //先找正上方有没有站名，没有的话向右找
                                    if (StationRow.GetCell(trainColumn) != null &&
                                        StationRow.GetCell(trainColumn).ToString().Trim().Length != 0)
                                    {
                                        _s.stationName = StationRow.GetCell(trainColumn).ToString().Trim();
                                    }
                                    //+1是“发车”
                                    if (TitleRow.GetCell(trainColumn + 1) != null)
                                    {
                                        if (TitleRow.GetCell(trainColumn + 1).ToString().Trim().Equals("发车"))
                                        {
                                            //如果站名还没填上，则填上
                                            if (_s.stationName.Length == 0 && StationRow.GetCell(trainColumn + 1) != null &&
                                            StationRow.GetCell(trainColumn + 1).ToString().Trim().Length != 0)
                                            {
                                                _s.stationName = StationRow.GetCell(trainColumn + 1).ToString().Trim();
                                            }
                                            //填上发车点
                                            if (trainRow.GetCell(trainColumn + 1) != null && trainRow.GetCell(trainColumn + 1).ToString().Trim().Length != 0)
                                            {
                                                _s.startedTime = trainRow.GetCell(trainColumn + 1).ToString().Trim();
                                            }
                                        }
                                    }
                                    //如果站名还没找到，向左找
                                    //没有找到站名，向左找
                                    if (_s.stationName.Length == 0)
                                    {
                                            if (TitleRow.GetCell(trainColumn - 1).ToString().Trim().Equals("到达"))
                                            {
                                                if (StationRow.GetCell(trainColumn - 1) != null &&
                                                    StationRow.GetCell(trainColumn - 1).ToString().Trim().Length != 0)
                                                {
                                                    _s.stationName = StationRow.GetCell(trainColumn - 1).ToString().Trim();
                                                }
                                            }
                                    }
                                }
                                else if (TitleRow.GetCell(trainColumn).ToString().Trim().Equals("发车"))
                                {//如果这一行左边有字，且左边是股道，则添加过车站了，跳过
                                    if (trainRow.GetCell(trainColumn - 1) != null&&
                                        trainRow.GetCell(trainColumn - 1).ToString().Trim().Length != 0 &&
                                            TitleRow.GetCell(trainColumn - 1) != null &&
                                        TitleRow.GetCell(trainColumn - 1).ToString().Trim().Length != 0)
                                    {
                                        if (TitleRow.GetCell(trainColumn - 1).ToString().Trim().Equals("股道"))
                                        {
                                            continue;
                                        }
                                    }
                                    //把发车点填上
                                    _s.startedTime = trainRow.GetCell(trainColumn).ToString().Trim();
                                    //如果没有站名，尝试在上面找站名
                                    if (_s.stationName.Length == 0 &&
                                        StationRow.GetCell(trainColumn) != null &&
                                       StationRow.GetCell(trainColumn).ToString().Trim().Length != 0)
                                    {
                                        _s.stationName = StationRow.GetCell(trainColumn).ToString().Trim();
                                    }
                                    //没有找到站名，向左找
                                    if(_s.stationName.Length == 0)
                                    {
                                            if (TitleRow.GetCell(trainColumn - 1).ToString().Trim().Equals("股道"))
                                            {
                                                if (StationRow.GetCell(trainColumn - 1) != null &&
                                                    StationRow.GetCell(trainColumn - 1).ToString().Trim().Length != 0)
                                                {
                                                    _s.stationName = StationRow.GetCell(trainColumn-1).ToString().Trim();
                                                }
                                            }
                                    }
                                    //还没有，继续向左
                                    if (_s.stationName.Length == 0)
                                    {

                                            if (TitleRow.GetCell(trainColumn - 2).ToString().Trim().Equals("到达"))
                                            {
                                                if (StationRow.GetCell(trainColumn - 2) != null &&
                                                    StationRow.GetCell(trainColumn - 2).ToString().Trim().Length != 0)
                                                {
                                                    _s.stationName = StationRow.GetCell(trainColumn - 2).ToString().Trim();
                                                }
                                            }
                                    }

                                }
                                //查重
                                bool hasGotIt = false;
                                foreach(Station _station in _train.newStations)
                                {
                                    if (_station.stationName.Trim().Equals(_s.stationName.Trim()))
                                    {
                                        hasGotIt = true;
                                    }
                                }
                                if (!hasGotIt)
                                {
                                    _train.newStations.Add(_s);
                                }

                            }
                            else
                            {
                                continue;
                            }
                            
                        }
                        allTrains.Add(_train);
                    }
                }
                else
                {
                    MessageBox.Show("请将临客表内包含需要导出列车的表格单独复制至新Excel文件内，新汇总时刻表没有“车次”，“始发”，“终到”，“行别”列，请检查", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
             return allTrains;
        }

        private List<Train> GetTrainsFromNewTimeTables()
        {
            //对于每一个工作簿，先把左边一列的列位置找出，然后在时刻表中根据行来确定车站名称
            //发现“车次”字样后，右边的都是车次，根据车次所在位置往上找，(为空的向左上找)若左边对应列为“终到”，则为终到站，若为“始发”，则为始发站。
            //双车次在检测到车次的时候就进行分离，第一车次需要和上/下行对应（寻找周边的车次）
            //找到每一个车次后，直接对该车次的时刻表/股道进行添加，若往右已经没有了则结束。
            //当找到下一个“始发站”的时候，意味着是下一组车次。
            List<Train> trains = new List<Train>();
            foreach (IWorkbook workbook in NewTimeTablesWorkbooks)
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {//获取所有工作表
                    ISheet sheet = workbook.GetSheetAt(i);
                    IRow row;
                    //表头数据
                    int[] _startStationRaw = new int[50];
                    int[] _stopStationRaw = new int[50];
                    int trainRawCounter = 0;
                    int[] _trainRawNum = new int[50];
                    int titleColumn = 0;
                    //已经找到，不再继续找
                    //bool shouldContinue = true;
                    //上行双数false 下行单数true
                    bool[] upOrDown = new bool[50];
                    for (int j = 0; j < sheet.LastRowNum; j++)
                    {//找表头数据
                        row = sheet.GetRow(j);
                        if (row != null)
                        {
                            for (int k = 0; k < row.LastCellNum; k++)
                            {
                                if (row.GetCell(k) != null)
                                {
                                    //先确定标题
                                    if (row.GetCell(k).ToString().Trim().Contains("始发"))
                                    {//始发站所在行
                                        _startStationRaw[trainRawCounter] = j;
                                        //标题所在列
                                        titleColumn = k;
                                        break;
                                    }
                                    if (row.GetCell(k).ToString().Trim().Contains("终到"))
                                    {//终到站所在行
                                        _stopStationRaw[trainRawCounter] = j;
                                        break;
                                    }
                                    if (row.GetCell(k).ToString().Trim().Contains("车次"))
                                    {//车次所在行
                                        _trainRawNum[trainRawCounter] = j;
                                        trainRawCounter++;
                                        //shouldContinue = false;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //根据表头数据找该组车次
                    // stations_tb.Text = "始发行：" + startStationRaw + " 终到行：" + stopStationRaw + " 车次行：" + trainNumRaw + " 标题列：" + titleColumn;

                    //找当前sheet的某一组车是上行还是下行=。=
                    //逻辑：先找到一个车次，再往下找格子，如果格子是“时刻”，就继续往下找格子，如果有冒号，中文“：”换成英文，
                    //有冒号的格子存储一下，再往下直到找到第二个带冒号的格子
                    //去掉冒号，若是六位数字，只取前四位。
                    //去掉冒号后进行对比大小，若1<2，则为下行-单数-true，1>2则为上行-双数-false
                    int _trainRowNumLen = 0;
                    for (; _trainRowNumLen < _trainRawNum.Length; _trainRowNumLen++)
                    {
                        if (_trainRowNumLen != 0 && _trainRawNum[_trainRowNumLen] == 0)
                        {
                            break;
                        }
                    }
                    //用于某种特殊情况：带小时的没找够两个，但是带分钟的找够了两个，此时不能向右转移到下一个车次上去。
                    bool continueSearch = true;
                    //找每一个车次行是上行还是下行
                    for (int m = 0; m < _trainRowNumLen; m++)
                    {//找每一个车次行是上行还是下行
                        IRow trainRow = sheet.GetRow(_trainRawNum[m]);
                        //判断有没有找齐两个时间
                        int _continueCounter_hour = 0;
                        int _continueCounter_minute = 0;
                        //两个时间，如果有小时的话优先用小时，否则用分钟对比
                        int[] _trainTimeWithHour = new int[2];
                        int[] _trainTimeWithMinute = new int[2];
                        /*不判断上下行
                        for (int t = titleColumn + 1; t < trainRow.LastCellNum; t++)
                        {//t为列，firstTrainRaw为行
                            _trainTimeWithMinute = new int[2];
                            _trainTimeWithHour = new int[2];
                            _continueCounter_hour = 0;
                            _continueCounter_minute = 0;
                            if (trainRow.GetCell(t) != null)
                            {//开始往下找了
                                if (trainRow.GetCell(t).ToString().Trim().Length != 0)
                                {
                                    //读取的当前组的终点
                                    int _loadedLastRaw = 0;
                                    if (m == _trainRowNumLen - 1)
                                    {
                                        _loadedLastRaw = sheet.LastRowNum;
                                    }
                                    else if (m < _trainRowNumLen - 1)
                                    {
                                        _loadedLastRaw = _startStationRaw[m + 1];
                                    }
                                    else
                                    {
                                        return;
                                    }
                                    for (int tt = _trainRawNum[m]; tt <= _loadedLastRaw; tt++)
                                    {
                                        IRow tempRaw = sheet.GetRow(tt);
                                        string cellInfo = "";
                                        if (tempRaw != null)
                                            if (tempRaw.GetCell(t) != null)
                                            {
                                                cellInfo = tempRaw.GetCell(t).ToString();
                                                if (cellInfo.Contains(":") ||
                                                    cellInfo.Contains("："))
                                                {
                                                    //只要找到时间了，就不能换列了
                                                    continueSearch = false;
                                                    cellInfo = cellInfo.Replace("：", "").Trim();
                                                    cellInfo = cellInfo.Replace(":", "").Trim();
                                                    int _foundTime = 0;
                                                    if (cellInfo.Length % 2 == 0)
                                                    {//是四位数字或者六位数字
                                                        int.TryParse(cellInfo.Substring(0, 4), out _foundTime);
                                                    }
                                                    else
                                                    {
                                                        int.TryParse(cellInfo.Substring(0, 3), out _foundTime);
                                                    }
                                                    if (_foundTime != 0)
                                                    {
                                                        _trainTimeWithHour[_continueCounter_hour] = _foundTime;
                                                        _continueCounter_hour++;
                                                    }
                                                }
                                                else if (cellInfo.Trim().Length > 0 &&
                                                    _continueCounter_minute < 2)
                                                {//把只有分钟的也存一组
                                                    cellInfo = cellInfo.Trim();
                                                    Regex r = new Regex(@"^[0-9]+$");
                                                    if (r.Match(cellInfo).Success)
                                                    {//仅包含数字
                                                        continueSearch = false;
                                                        int _foundTime = 0;
                                                        if (cellInfo.Length % 2 == 0)
                                                        {//是四位数字或者六位数字
                                                            int.TryParse(cellInfo.Substring(0, 2), out _foundTime);
                                                        }
                                                        else
                                                        {
                                                            int.TryParse(cellInfo.Substring(0, 1), out _foundTime);
                                                        }
                                                        if (_foundTime != 0)
                                                        {
                                                            _trainTimeWithMinute[_continueCounter_minute] = _foundTime;
                                                            _continueCounter_minute++;
                                                        }
                                                    }
                                                }
                                                if (_continueCounter_hour == 2)
                                                {
                                                    break;
                                                }
                                            }
                                    }
                                    if (_trainTimeWithMinute[0] == 0 && _trainTimeWithMinute[1] == 0 && _continueCounter_hour != 2)
                                    {
                                        continueSearch = true;
                                    }

                                }
                            }
                            //如果找到了一个小时一个分钟 必须重找
                            if (_continueCounter_hour == 1 && _continueCounter_minute == 1)
                            {
                                continueSearch = true;
                            }
                            if (!continueSearch)
                            {
                                break;
                            }
                        }
                        if (_continueCounter_hour != 2)
                        {//小时没取到，就取分钟计算吧
                            if (_trainTimeWithMinute[0] < _trainTimeWithMinute[1])
                            {//下行
                                upOrDown[m] = true;
                            }
                            else if (_trainTimeWithMinute[0] > _trainTimeWithMinute[1])
                            {//上行
                                upOrDown[m] = false;
                            }
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            //如果一个23点，一个0点，（A-B>=20或B-A>=20的时候）要单独判断
                           // if (_trainTimeWithHour[0] - _trainTimeWithHour[1] > 20)
                            //{
                           //     upOrDown[m] = true;
                           // }
                           // else if (_trainTimeWithHour[1] - _trainTimeWithHour[0] > 20)
                           // {
                            //    upOrDown[m] = false;
                           // }
                            //
                            if (_trainTimeWithHour[0] < _trainTimeWithHour[1])
                            {//下行
                                upOrDown[m] = true;
                                if (_trainTimeWithHour[0] / 100 == 0 && _trainTimeWithHour[1] / 100 == 23)
                                {
                                    upOrDown[m] = false;
                                }
                            }
                            else if (_trainTimeWithHour[0] > _trainTimeWithHour[1])
                            {//上行
                                upOrDown[m] = false;
                                if (_trainTimeWithHour[0] / 100 == 23 && _trainTimeWithHour[1] == 0)
                                {
                                    upOrDown[m] = true;
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                    */
                    }


                    //=========
                    //开始添加该sheet中的车次
                    //=========
                    for (int _rowNum = 0; _rowNum < _trainRowNumLen; _rowNum++)
                    {
                        if (_trainRawNum[_rowNum] == 0)
                        {//如果车次所在行号是0的话就不搜索，否则进行搜索
                            continue;
                        }
                        else
                        {
                            //从第一个有车次的行开始找
                            IRow tempRow = sheet.GetRow(_trainRawNum[_rowNum]);
                            for (int t = titleColumn + 1; t < tempRow.LastCellNum; t++)
                            {//从第一列开始找
                                if (tempRow != null)
                                {
                                    if (tempRow.GetCell(t) != null)
                                    {
                                        if (tempRow.GetCell(t).ToString().Trim().Contains("G") ||
                                            tempRow.GetCell(t).ToString().Trim().Contains("D") ||
                                            tempRow.GetCell(t).ToString().Trim().Contains("C") ||
                                            tempRow.GetCell(t).ToString().Trim().Contains("J"))
                                        {
                                            //车次模型
                                            Train tempTrain = new Train();
                                            //==============
                                            //找起点站终点站
                                            //找法：
                                            //找到车次后直接往终到站那行-车次列找-如果没有，就往左一格一格找（找到的字不能是终到站，如果没有了就放弃）
                                            //找始发站往始发站行-车次列找，没有的往左一格一格找
                                            IRow stopRow = sheet.GetRow(_stopStationRaw[_rowNum]);
                                            IRow startRow = sheet.GetRow(_startStationRaw[_rowNum]);
                                            //找始发站-终到站（一起找吧…）

                                            bool continueFindingStart = true;
                                            bool continueFindingStop = true;
                                            for (int s = t; s > titleColumn; s--)
                                            {
                                                string startStation = "";
                                                string stopStation = "";
                                                if (stopRow.GetCell(s) != null)
                                                {
                                                    stopStation = stopRow.GetCell(s).ToString().Trim();
                                                    if (!stopStation.Equals("") &&
                                                        !stopStation.Contains("终到") &&
                                                        continueFindingStop)
                                                    {
                                                        if (stopStation.Trim().Contains("东动车所"))
                                                        {
                                                            stopStation = "郑州东动车所";
                                                        }
                                                        if (stopStation.Trim().Contains("南动车所"))
                                                        {
                                                            stopStation = "郑州南动车所";
                                                        }
                                                        if (stopStation.Trim().Contains("焦作"))
                                                        {
                                                            stopStation = "焦作";
                                                        }
                                                        tempTrain.stopStation = stopStation.Trim();
                                                        continueFindingStop = false;
                                                    }
                                                }
                                                if (startRow.GetCell(s) != null)
                                                {
                                                    startStation = startRow.GetCell(s).ToString().Trim();
                                                    if (!startStation.Equals("") &&
                                                        !startStation.Contains("始发") &&
                                                        continueFindingStart)
                                                    {
                                                        if (stopStation.Trim().Contains("东动车所"))
                                                        {
                                                            stopStation = "郑州东动车所";
                                                        }
                                                        if (stopStation.Trim().Contains("南动车所"))
                                                        {
                                                            stopStation = "郑州南动车所";
                                                        }
                                                        if (startStation.Trim().Contains("焦作"))
                                                        {
                                                            startStation = "焦作";
                                                        }
                                                        tempTrain.startStation = startStation.Trim();
                                                        continueFindingStart = false;
                                                    }
                                                }
                                            }

                                            //===============
                                            if (tempRow.GetCell(t).ToString().Trim().Contains("/"))
                                            {//双车次
                                                string[] TrainNums = splitTrainNum(tempRow.GetCell(t).ToString().Trim());
                                                tempTrain.firstTrainNum = TrainNums[0];
                                                tempTrain.secondTrainNum = TrainNums[1];
                                            }
                                            else
                                            {
                                                tempTrain.firstTrainNum = tempRow.GetCell(t).ToString().Trim();
                                            }

                                            //数据为某一组的上下行，下行单号true，上行双号false
                                            //tempTrain.upOrDown = upOrDown[_rowNum];
                                            //取末位数字模2
                                            int number = 0;
                                            string num = null;
                                            foreach (char item in tempTrain.firstTrainNum)
                                            {
                                                if (item >= 48 && item <= 58)
                                                {
                                                    num += item;
                                                }
                                            }
                                            number = int.Parse(num);
                                            if(number%2 == 0)
                                            {
                                                tempTrain.upOrDown = false ;
                                            }
                                            else
                                            {
                                                tempTrain.upOrDown = true;
                                            }
                                            //找停的车站
                                            List<Station> tempStations = new List<Station>();

                                            //往下找时刻和股道
                                            //找法：
                                            //找到车次后，根据上下行的情况，下行则↓搜索，上行则↑搜索。
                                            //找到有内容的格子后：
                                            //根据这个时间能不能在左侧找到站名来判断是什么时间
                                            //↑ 发√  到
                                            //↓ 发      到√
                                            //如果直接往左找不到，就左上找
                                            //如果找到的是带小时的时间格子，则将“：”分隔开的左半部分当成小时，在找到新的小时标记之前，沿用这个小时标记。
                                            //如果找到的第一个时间是发车时间，则：将该站的停站模式改为1，同时新建模型
                                            //0-普通-有停有发，1-始发-有被接续有发，2-终到-有到有被接续，3-通过
                                            //若股道为罗马数字 则转换为阿拉伯数字
                                            //满足一组到-发后，即可建立一个车站模型
                                            int stoppedRow = 0;
                                            if (_rowNum == _trainRowNumLen - 1)
                                            {//判断搜索的下边界
                                                stoppedRow = sheet.LastRowNum;
                                            }
                                            else
                                            {
                                                stoppedRow = _startStationRaw[_rowNum + 1];
                                            }
                                            //根据上下行找车次
                                            //if (upOrDown[_rowNum]){
                                           //下行
                                                //下行
                                                //下行
                                                //下行
                                                //下行
                                                //下行
                                                //下行
                                                //下行
                                                //下行
                                                //下行
                                                //下行
                                                //用完一组之后重置
                                                string _hours = "";
                                                string _tempStoppedTime = "";
                                                string _tempStartingTime = "";
                                                string _stationName = "";
                                                int _stationType = 0;
                                                string _track = "";
                                                for (int tt = _trainRawNum[_rowNum]; tt <= stoppedRow; tt++)
                                                {
                                                    string _foundTime = "";
                                                    IRow _trainTimeRow = sheet.GetRow(tt);
                                                    if (_trainTimeRow == null)
                                                    {
                                                        continue;
                                                    }
                                                    if (_trainTimeRow.GetCell(t) != null)
                                                    {
                                                        //找时间
                                                        string cellInfo = _trainTimeRow.GetCell(t).ToString();
                                                        if (cellInfo.Contains(":") ||
                                                            cellInfo.Contains("："))
                                                        {
                                                            cellInfo = cellInfo.Replace("：", ":").Trim();
                                                            //有小时的，把小时存储起来，在找到下一个小时前继续用
                                                            _hours = cellInfo.Split(':')[0];
                                                            cellInfo = cellInfo.Replace(":", "").Trim();
                                                            if (cellInfo.Length % 2 == 0)
                                                            {//是四位数字或者六位数字，此处取分钟
                                                                _foundTime = cellInfo.Substring(2, 2);
                                                            }
                                                            else
                                                            {
                                                                _foundTime = cellInfo.Substring(1, 2);
                                                            }
                                                            _foundTime = _hours + ":" + _foundTime;
                                                        }
                                                        else if (cellInfo.Contains(".") ||
                                                            cellInfo.Contains("…"))
                                                        {//通过车
                                                            _foundTime = "通过";
                                                            _stationType = 3;
                                                        }
                                                        else if (cellInfo.Contains("-"))
                                                        {//终到了
                                                            _stationType = 2;
                                                            _foundTime = "终到";
                                                        }
                                                        else if (cellInfo.Trim().Length > 0)
                                                        {//把只有分钟的也存一组
                                                            cellInfo = cellInfo.Trim();
                                                            Regex r = new Regex(@"^[0-9]+$");
                                                            if (r.Match(cellInfo).Success)
                                                            {//仅包含数字
                                                                if (cellInfo.Length % 2 == 0)
                                                                {//是四位数字或者六位数字
                                                                    _foundTime = cellInfo.Substring(0, 2);
                                                                }
                                                                else
                                                                {
                                                                    _foundTime = cellInfo.Substring(0, 1);
                                                                }
                                                            }
                                                            if (cellInfo.Trim().Length > 0 &&
                                                                !_hours.Equals("") &&
                                                                r.Match(cellInfo).Success)
                                                            {
                                                                _foundTime = _hours + ":" + _foundTime;
                                                            }
                                                        }

                                                        //判断是到还是发-检查左边的格子有没有车站名-检查右边的格子有没有股道号
                                                        //如果↓第一个时间是发点，就获取不到数据，此时应当获取上一行的数据
                                                        string tempStation = "";
                                                        if (_trainTimeRow.GetCell(titleColumn) != null)
                                                        {//找站名      
                                                            tempStation = _trainTimeRow.GetCell(titleColumn).ToString().Trim();
                                                            if (tempStation.Length != 0)
                                                            {
                                                                _stationName = tempStation;
                                                                if (tempStation.Contains("焦作"))
                                                                {
                                                                    _stationName = "焦作";
                                                                }
                                                                //是到点
                                                                _tempStoppedTime = _foundTime;
                                                            }
                                                            else
                                                            {
                                                                if (tt > 0)
                                                                {//从上一行找
                                                                    if (sheet.GetRow(tt - 1) != null)
                                                                    {
                                                                        IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                        if (lastTrainRow != null)
                                                                        {
                                                                            if (lastTrainRow.GetCell(titleColumn) != null)
                                                                            {
                                                                                tempStation = lastTrainRow.GetCell(titleColumn).ToString().Trim();
                                                                                if (tempStation.Length != 0)
                                                                                {
                                                                                    _stationName = tempStation;
                                                                                    if (tempStation.Contains("焦作"))
                                                                                    {
                                                                                        _stationName = "焦作";
                                                                                    }
                                                                                    //是发点
                                                                                    _tempStartingTime = _foundTime;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                        }
                                                        else
                                                        {
                                                            if (tt > 0)
                                                            {//从上一行找
                                                                if (sheet.GetRow(tt - 1) != null)
                                                                {
                                                                    IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                    if (lastTrainRow != null)
                                                                    {
                                                                        if (lastTrainRow.GetCell(titleColumn) != null)
                                                                        {
                                                                            tempStation = lastTrainRow.GetCell(titleColumn).ToString().Trim();
                                                                            if (tempStation.Length != 0)
                                                                            {
                                                                                _stationName = tempStation;
                                                                                if (tempStation.Contains("焦作"))
                                                                                {
                                                                                    _stationName = "焦作";
                                                                                }
                                                                                //是发点
                                                                                _tempStartingTime = _foundTime;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if (!_foundTime.Equals(_tempStartingTime) &&
                                                           !_foundTime.Equals(_tempStartingTime))
                                                        {//如果时间没有填进去？？

                                                        }
                                                        string tempTrack = "";
                                                        if (t < tempRow.LastCellNum - 1)
                                                        {
                                                            if (_trainTimeRow.GetCell(t + 1) != null)
                                                                if (!_trainTimeRow.GetCell(t).ToString().Trim().Equals(""))
                                                                {//找股道
                                                                    tempTrack = _trainTimeRow.GetCell(t + 1).ToString().Trim();
                                                                    if (tempTrack.Length != 0)
                                                                    {
                                                                        _track = tempTrack;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (tt > 0)
                                                                        {//从上一行找
                                                                            if (sheet.GetRow(tt - 1) != null)
                                                                            {
                                                                                IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                                if (lastTrainRow != null)
                                                                                {
                                                                                    if (lastTrainRow.GetCell(t + 1) != null)
                                                                                    {
                                                                                        tempTrack = lastTrainRow.GetCell(t + 1).ToString().Trim();
                                                                                        if (tempTrack.Length != 0)
                                                                                        {
                                                                                            _track = tempTrack;
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                        }
                                                        else
                                                        {
                                                            if (tt > 0)
                                                            {//从上一行找
                                                                if (sheet.GetRow(tt - 1) != null)
                                                                {
                                                                    IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                    if (lastTrainRow != null)
                                                                    {
                                                                        if (lastTrainRow.GetCell(t + 1) != null)
                                                                        {
                                                                            tempTrack = lastTrainRow.GetCell(t + 1).ToString().Trim();
                                                                            if (tempTrack.Length != 0)
                                                                            {
                                                                                _track = tempTrack;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        Station _tempStation = new Station();
                                                        if (_tempStartingTime.Length != 0 &&
                                                            _tempStoppedTime.Length != 0)
                                                        {//已经获取到一组数据，此时应当添加模型
                                                            if (_track.Equals("XXV"))
                                                            {
                                                                _track = "25";
                                                            }
                                                            else if (_track.Equals("XXVI"))
                                                            {
                                                                _track = "26";
                                                            }
                                                            _tempStation.stoppedTime = _tempStoppedTime;
                                                            _tempStation.startedTime = _tempStartingTime;
                                                            _tempStation.stationType = _stationType;
                                                            _tempStation.stationName = _stationName;
                                                            _tempStation.stationTrackNum = _track;
                                                            //此处修改添加新主站（南站模式）
                                                            if (_stationName.Contains("郑州东") &&
                                                                 !_stationName.Contains("动车所") &&
                                                                 !_stationName.Contains("疏解区") &&
                                                                 !_stationName.Contains("郑州南"))
                                                            {
                                                                if (_tempStartingTime.Contains("终"))
                                                                {
                                                                    _tempStation.startedTime = _tempStoppedTime.Replace(":", "");
                                                                    //_tempStation.startedTime = _tempStoppedTime;
                                                                }
                                                                else
                                                                {
                                                                    _tempStation.startedTime = _tempStartingTime.Replace(":", "");
                                                                    //_tempStation.startedTime = _tempStartingTime;
                                                                }
                                                                tempTrain.mainStation = _tempStation;
                                                            }
                                                            else
                                                            {
                                                                tempStations.Add(_tempStation);
                                                            }
                                                            //清零
                                                            _tempStartingTime = "";
                                                            _tempStoppedTime = "";
                                                            _stationType = 0;
                                                            _stationName = "";
                                                            _track = "";
                                                        }
                                                        else if (_tempStartingTime.Length != 0 &&
                                                            _tempStoppedTime.Length == 0)
                                                        {//只有发时，没有停时，此站为始发站
                                                            if (_track.Equals("XXV"))
                                                            {
                                                                _track = "25";
                                                            }
                                                            else if (_track.Equals("XXVI"))
                                                            {
                                                                _track = "26";
                                                            }
                                                            _tempStation.startedTime = _tempStartingTime;
                                                            _tempStation.stoppedTime = "始发";
                                                            _tempStation.stationType = 1;
                                                            _tempStation.stationName = _stationName;
                                                            _tempStation.stationTrackNum = _track;
                                                            //此处修改添加新主站（南站模式）
                                                            if (_stationName.Contains("郑州东") &&
                                                                 !_stationName.Contains("动车所") &&
                                                                 !_stationName.Contains("疏解区") &&
                                                                 !_stationName.Contains("郑州南"))
                                                            {
                                                                _tempStation.startedTime = _tempStartingTime.Replace(":", "");
                                                                //_tempStation.startedTime = _tempStartingTime;
                                                                tempTrain.mainStation = _tempStation;
                                                            }
                                                            else
                                                            {
                                                                tempStations.Add(_tempStation);
                                                            }
                                                            //清零
                                                            _tempStartingTime = "";
                                                            _tempStoppedTime = "";
                                                            _stationType = 0;
                                                            _stationName = "";
                                                            _track = "";
                                                        }
                                                    }
                                                }
                                                /*
                                            }
                                            
                                            else
                                            {//上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //上行
                                             //用完一组之后重置
                                                string _hours = "";
                                                string _tempStoppedTime = "";
                                                string _tempStartingTime = "";
                                                string _stationName = "";
                                                int _stationType = 0;
                                                string _track = "";
                                                //判断一下是不是始发站
                                                bool onlyStart = false;
                                                for (int tt = stoppedRow; tt > _trainRawNum[_rowNum]; tt--)
                                                {
                                                    string _foundTime = "";
                                                    IRow _trainTimeRow = sheet.GetRow(tt);
                                                    if (_trainTimeRow != null)
                                                        if (_trainTimeRow.GetCell(t) != null)
                                                        {
                                                            //找时间
                                                            string cellInfo = _trainTimeRow.GetCell(t).ToString();
                                                            if (cellInfo.Contains(":") ||
                                                                cellInfo.Contains("："))
                                                            {
                                                                cellInfo = cellInfo.Replace("：", ":").Trim();
                                                                //有小时的，把小时存储起来，在找到下一个小时前继续用
                                                                _hours = cellInfo.Split(':')[0];
                                                                cellInfo = cellInfo.Replace(":", "").Trim();
                                                                if (cellInfo.Length % 2 == 0)
                                                                {//是四位数字或者六位数字，此处取分钟
                                                                    _foundTime = cellInfo.Substring(2, 2);
                                                                }
                                                                else
                                                                {
                                                                    _foundTime = cellInfo.Substring(1, 2);
                                                                }
                                                                _foundTime = _hours + ":" + _foundTime;
                                                            }
                                                            else if (cellInfo.Contains(".") ||
                                                                cellInfo.Contains("…"))
                                                            {//通过车
                                                                _foundTime = "通过";
                                                                _stationType = 3;
                                                            }
                                                            else if (cellInfo.Contains("-"))
                                                            {//终到了
                                                                _stationType = 2;
                                                                _foundTime = "终到";
                                                            }
                                                            else if (cellInfo.Trim().Length > 0)
                                                            {//把只有分钟的也存一组
                                                                cellInfo = cellInfo.Trim();
                                                                Regex r = new Regex(@"^[0-9]+$");
                                                                if (r.Match(cellInfo).Success)
                                                                {//仅包含数字
                                                                    if (cellInfo.Length % 2 == 0)
                                                                    {//是四位数字或者六位数字
                                                                        _foundTime = cellInfo.Substring(0, 2);
                                                                    }
                                                                    else
                                                                    {
                                                                        _foundTime = cellInfo.Substring(0, 1);
                                                                    }
                                                                }
                                                                if (cellInfo.Trim().Length > 0 &&
                                                                        !_hours.Equals("") &&
                                                                        r.Match(cellInfo).Success)
                                                                {
                                                                    _foundTime = _hours + ":" + _foundTime;
                                                                }

                                                            }

                                                            //判断是到还是发-检查左边的格子有没有车站名-检查右边的格子有没有股道号
                                                            //如果↓第一个时间是发点，就获取不到数据，此时应当获取上一行的数据
                                                            string tempStation = "";
                                                            if (_trainTimeRow.GetCell(titleColumn) != null)
                                                            {//找站名      
                                                                tempStation = _trainTimeRow.GetCell(titleColumn).ToString().Trim();
                                                                if (tempStation.Length != 0)
                                                                {
                                                                    _stationName = tempStation;
                                                                    if (tempStation.Contains("焦作"))
                                                                    {
                                                                        _stationName = "焦作";
                                                                    }
                                                                    //是发点
                                                                    _tempStartingTime = _foundTime;
                                                                }
                                                                else
                                                                {
                                                                    if (tt > 0)
                                                                    {//从上一行找
                                                                        if (sheet.GetRow(tt - 1) != null)
                                                                        {
                                                                            IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                            if (lastTrainRow != null)
                                                                            {
                                                                                if (lastTrainRow.GetCell(titleColumn) != null)
                                                                                {
                                                                                    tempStation = lastTrainRow.GetCell(titleColumn).ToString().Trim();
                                                                                    if (tempStation.Length != 0)
                                                                                    {
                                                                                        _stationName = tempStation;
                                                                                        if (tempStation.Contains("焦作"))
                                                                                        {
                                                                                            _stationName = "焦作";
                                                                                        }
                                                                                        //是到点
                                                                                        _tempStoppedTime = _foundTime;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                            }
                                                            else
                                                            {
                                                                if (tt > 0)
                                                                {//从上一行找
                                                                    if (sheet.GetRow(tt - 1) != null)
                                                                    {
                                                                        IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                        if (lastTrainRow != null)
                                                                        {
                                                                            if (lastTrainRow.GetCell(titleColumn) != null)
                                                                            {
                                                                                tempStation = lastTrainRow.GetCell(titleColumn).ToString().Trim();
                                                                                if (tempStation.Length != 0)
                                                                                {
                                                                                    _stationName = tempStation;
                                                                                    if (tempStation.Contains("焦作"))
                                                                                    {
                                                                                        _stationName = "焦作";
                                                                                    }
                                                                                    //是到点
                                                                                    _tempStoppedTime = _foundTime;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            if (!_foundTime.Equals(_tempStartingTime) &&
                                                               !_foundTime.Equals(_tempStartingTime))
                                                            {//如果时间没有填进去？？

                                                            }
                                                            string tempTrack = "";
                                                            if (t < tempRow.LastCellNum - 1)
                                                            {
                                                                if (_trainTimeRow.GetCell(t + 1) != null)
                                                                {//找股道
                                                                    tempTrack = _trainTimeRow.GetCell(t + 1).ToString().Trim();
                                                                    if (tempTrack.Length != 0)
                                                                    {
                                                                        _track = tempTrack;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (tt > 0)
                                                                        {//从上一行找
                                                                            if (sheet.GetRow(tt - 1) != null)
                                                                            {
                                                                                IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                                if (lastTrainRow != null)
                                                                                {
                                                                                    if (lastTrainRow.GetCell(t + 1) != null)
                                                                                    {
                                                                                        tempTrack = lastTrainRow.GetCell(t + 1).ToString().Trim();
                                                                                        if (tempTrack.Length != 0)
                                                                                        {
                                                                                            _track = tempTrack;
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (tt > 0)
                                                                {//从上一行找
                                                                    if (sheet.GetRow(tt - 1) != null)
                                                                    {
                                                                        IRow lastTrainRow = sheet.GetRow(tt - 1);
                                                                        if (lastTrainRow != null)
                                                                        {
                                                                            if (lastTrainRow.GetCell(t + 1) != null)
                                                                            {
                                                                                tempTrack = lastTrainRow.GetCell(t + 1).ToString().Trim();
                                                                                if (tempTrack.Length != 0)
                                                                                {
                                                                                    _track = tempTrack;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            Station _tempStation = new Station();
                                                            if (_tempStartingTime.Length != 0 &&
                                                                _tempStoppedTime.Length != 0)
                                                            {//已经获取到一组数据，此时应当添加模型
                                                                if (_track.Equals("XXV"))
                                                                {
                                                                    _track = "25";
                                                                }
                                                                else if (_track.Equals("XXVI"))
                                                                {
                                                                    _track = "26";
                                                                }
                                                                _tempStation.stoppedTime = _tempStoppedTime;
                                                                _tempStation.startedTime = _tempStartingTime;
                                                                _tempStation.stationType = _stationType;
                                                                _tempStation.stationName = _stationName;
                                                                _tempStation.stationTrackNum = _track;
                                                                //添加南站修改处
                                                                if (_stationName.Contains("郑州东") &&
                                                                     !_stationName.Contains("动车所") &&
                                                                     !_stationName.Contains("疏解区") &&
                                                                     !_stationName.Contains("郑州南"))
                                                                {
                                                                    if (_tempStartingTime.Contains("终"))
                                                                    {
                                                                        _tempStation.startedTime = _tempStoppedTime.Replace(":", "");
                                                                        //_tempStation.startedTime = _tempStoppedTime;
                                                                    }
                                                                    else
                                                                    {
                                                                        _tempStation.startedTime = _tempStartingTime.Replace(":", "");
                                                                        //_tempStation.startedTime = _tempStartingTime;
                                                                    }
                                                                    tempTrain.mainStation = _tempStation;
                                                                }
                                                                else
                                                                {
                                                                    tempStations.Add(_tempStation);
                                                                }
                                                                //清零
                                                                _tempStartingTime = "";
                                                                _tempStoppedTime = "";
                                                                _stationType = 0;
                                                                _stationName = "";
                                                                _track = "";
                                                            }
                                                            else if (_tempStartingTime.Length != 0 &&
                                                                _tempStoppedTime.Length == 0)
                                                            {//只有发时，没有停时，此站为始发站
                                                                if (_track.Equals("XXV"))
                                                                {
                                                                    _track = "25";
                                                                }
                                                                else if (_track.Equals("XXVI"))
                                                                {
                                                                    _track = "26";
                                                                }
                                                                _tempStation.startedTime = _tempStartingTime;
                                                                _tempStation.stoppedTime = "始发";
                                                                _tempStation.stationType = 1;
                                                                _tempStation.stationName = _stationName;
                                                                _tempStation.stationTrackNum = _track;
                                                                //添加南站修改处
                                                                if (_stationName.Contains("郑州东") &&
                                                                    !_stationName.Contains("动车所") &&
                                                                    !_stationName.Contains("疏解区") &&
                                                                    !_stationName.Contains("郑州南"))
                                                                {
                                                                    _tempStation.startedTime = _tempStartingTime.Replace(":", "");
                                                                    //_tempStation.startedTime = _tempStartingTime;
                                                                    tempTrain.mainStation = _tempStation;
                                                                }
                                                                else
                                                                {
                                                                    tempStations.Add(_tempStation);
                                                                }
                                                                //清零
                                                                _tempStartingTime = "";
                                                                _tempStoppedTime = "";
                                                                _stationType = 0;
                                                                _stationName = "";
                                                                _track = "";
                                                            }
                                                        }
                                                }

                                            }
                                            */
                                                tempTrain.newStations = tempStations;
                                                if (tempTrain.firstTrainNum.Length != 0)
                                                {
                                                    trains.Add(tempTrain);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                }
                //allTrains_New = trains;

            }
            return trains;
        }



        private string[] splitTrainNum(string trainNum)
        {//分割车次
            string[] splitedTrainNum = new string[2];
            String[] trainWithDoubleNumber = trainNum.Split('/');
            //先添加第一个车次
            splitedTrainNum[0] = trainWithDoubleNumber[0];

            Char[] firstTrainWord = trainWithDoubleNumber[0].ToCharArray();
            String secondTrainWord = "";
            for (int q = 0; q < firstTrainWord.Length; q++)
            {
                if (q != firstTrainWord.Length - trainWithDoubleNumber[1].Length)
                {
                    secondTrainWord = secondTrainWord + firstTrainWord[q];
                }
                else
                {
                    secondTrainWord = secondTrainWord + trainWithDoubleNumber[1];
                    //添加第二个车次
                    splitedTrainNum[1] = secondTrainWord;
                    break;
                }
            }
            return splitedTrainNum;
        }

        //数据处理
        /*
        //重新修改文件指定单元格样式
        FileStream fs1 = File.OpenWrite(ExcelFile.FileName);
        workbook.Write(fs1);
         fs1.Close();
          Console.ReadLine();
          fileStream.Close();
          workbook.Close();

            数据处理
            1. 将车次分为有主站+无主站
            2.对于没有主站的车次，在全局中搜索其他车次寻找带主站的部分，将主站数据共享
            //如果主站为其他车复制的，则主站为“待定”，此时需要用自身车站与现有时刻表进行对比，
            直到发现A有B没有的站，或者B有A没有的站 才能确定属于哪张时刻表。
            3.对于徐兰场经过曹古寺的车，复制一份标注为京广场
            4.对于依然没有主站的车次（二郎庙-疏解区方向），判断它是否经过二郎庙和疏解区，是的话标注为京广场
            在后期分清上下行和场的时候插入进去排。
            5.将列车按照主站开车时刻进行排序（若主站无开车时间，以到达时间为基准）
            6.终到/始发列车寻找接续列车并标注-规则-在所有列车中寻找相同股道时间最近，并且满足相应匹配规则的
            例如 0G425 7:28到13道-无发出时间
            则匹配时间最近的在13道，且有发出无到达的列车，如果没有就不标注
            7.无主站列车，按照二郎庙通过时间排序。（插入式）
           */

        private List<Train> analyizeTrainData(List<Train> trains)
        {
            //有主站和没主站的-没主站的先从其他地方找相同车次
            List<Train> trainsWithMainStation = new List<Train>();
            List<Train> trainsWithoutMainStation = new List<Train>();
            //找主站
            foreach (Train train in trains)
            {
                if (train.mainStation != null && train.mainStation.stationName.Length != 0)
                {
                    trainsWithMainStation.Add(train);
                }
                else
                {
                    string firstTrainNumber = "";
                    string secondTrainNumber = "";
                    firstTrainNumber = train.firstTrainNum.Trim();
                    if (train.secondTrainNum != null && train.secondTrainNum.Length!= 0)
                    {
                        secondTrainNumber = train.secondTrainNum.Trim();
                    }
                    bool hasGotNumber = false;
                    foreach (Train tempTrain in trains)
                    {
                        if (tempTrain.firstTrainNum != null && tempTrain.firstTrainNum.Length!=0)
                            if (firstTrainNumber.Equals(tempTrain.firstTrainNum.Trim()) ||
                                    secondTrainNumber.Equals(tempTrain.firstTrainNum.Trim()))
                            {
                                if (tempTrain.mainStation != null && tempTrain.mainStation.stationName.Length != 0 && !hasGotNumber)
                                {
                                    Station tmp_Station = new Station();
                                    tmp_Station.stationName = "待定";
                                    tmp_Station.startedTime = tempTrain.mainStation.startedTime;
                                    tmp_Station.stoppedTime = tempTrain.mainStation.stoppedTime;
                                    tmp_Station.stationTrackNum = tempTrain.mainStation.stationTrackNum;
                                    train.mainStation = tmp_Station;
                                        hasGotNumber = true;
                                }
                                if (hasGotNumber)
                                {
                                    trainsWithMainStation.Add(train);
                                    break;
                                }
                            }
                            else if (tempTrain.secondTrainNum != null && tempTrain.secondTrainNum.Length != 0 && !hasGotNumber)
                            {
                                if (firstTrainNumber.Equals(tempTrain.secondTrainNum) ||
                                    secondTrainNumber.Equals(tempTrain.secondTrainNum))
                                {
                                    if (tempTrain.mainStation != null && tempTrain.mainStation.stationName.Length != 0)
                                    {
                                        Station tmp_Station = new Station();
                                        tmp_Station.stationName = "待定";
                                        tmp_Station.startedTime = tempTrain.mainStation.startedTime;
                                        tmp_Station.stoppedTime = tempTrain.mainStation.stoppedTime;
                                        tmp_Station.stationTrackNum = tempTrain.mainStation.stationTrackNum;
                                        train.mainStation = tmp_Station;
                                        hasGotNumber = true;
                                    }
                                }
                                if (hasGotNumber)
                                {
                                    trainsWithMainStation.Add(train);
                                    break;
                                }
                            }
                    }
                    if (!hasGotNumber)
                    {//如果都没有找到这个车次的话
                        trainsWithoutMainStation.Add(train);
                    }
                }
            }
            //寻找接续列车
            foreach (Train train in trainsWithMainStation)
            {//1始发2终到-由X改，续开X
                if (train.mainStation != null && train.mainStation.stationName.Length != 0)
                    if (train.mainStation.stationTrackNum != null && train.mainStation.stationTrackNum.Length != 0)
                        if (train.mainStation.stationType == 1)
                        {//找终到接续->由X改
                            Train tmp_Train = new Train();
                            foreach (Train compared_Train in trainsWithMainStation)
                            {
                                if (tmp_Train.mainStation == null || tmp_Train.mainStation.stationName.Length == 0)
                                {//首次寻找
                                    if (compared_Train.mainStation.stationTrackNum != null && compared_Train.mainStation.stationTrackNum.Length!=0)
                                    {
                                        if (compared_Train.mainStation.stationTrackNum.Equals(train.mainStation.stationTrackNum))
                                        {//找到了同股道列车-该列车必须终到
                                            if (compared_Train.mainStation.stationType != 2)
                                            {
                                                continue;
                                            }
                                            int trainTime = 0;
                                            int compared_TrainTime = 0;
                                            //比较-需要找的车的发车时间，和其他车的到达时间，且发车时间需要大于到达时间
                                            int.TryParse(train.mainStation.startedTime, out trainTime);
                                            int.TryParse(compared_Train.mainStation.stoppedTime.Replace(":", "").Trim(), out compared_TrainTime);

                                            if (trainTime > compared_TrainTime)
                                            {//新找到的时间和原时间差小于之前找到的时间差，用新的替换老的
                                                tmp_Train = compared_Train;
                                            }
                                        }
                                    }
                                }
                                else
                                {//非首次寻找，需要进行比较
                                    if (compared_Train.mainStation.stationTrackNum != null && compared_Train.mainStation.stationTrackNum.Length!=0)
                                    {
                                        if (compared_Train.mainStation.stationTrackNum.Equals(train.mainStation.stationTrackNum))
                                        {//找到了同股道列车-此车必须是终到车
                                            if(compared_Train.mainStation.stationType != 2)
                                            {
                                                continue;
                                            }
                                            int trainTime = 0;
                                            int compared_TrainTime = 0;
                                            int tmp_TrainTime = 0;
                                            //比较-需要找的车的发车时间，和其他车的到达时间，且发车时间需要大于到达时间
                                            int.TryParse(train.mainStation.startedTime, out trainTime);
                                            int.TryParse(compared_Train.mainStation.stoppedTime.Replace(":", "").Trim(), out compared_TrainTime);
                                            int.TryParse(tmp_Train.mainStation.stoppedTime.Replace(":", "").Trim(), out tmp_TrainTime);
                                            
                                            if ((Math.Abs(trainTime - compared_TrainTime) < Math.Abs(trainTime - tmp_TrainTime)) &&
                                                trainTime > compared_TrainTime)
                                            {//新找到的时间和原时间差小于之前找到的时间差，用新的替换老的
                                                tmp_Train = compared_Train;
                                            }
                                        }
                                    }
                                }
                                //如果有双车次，判断接续前列车的上下行
                                if (tmp_Train.secondTrainNum != null && tmp_Train.secondTrainNum.Length != 0)
                                {
                                    if (tmp_Train.upOrDown)
                                    {//下行
                                        Char[] TrainWord = tmp_Train.firstTrainNum.ToCharArray();
                                        if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                                        {//最后一位是偶数，则用奇数的那个
                                            train.mainStation.stoppedTime = "由" + tmp_Train.secondTrainNum + "改";
                                        }
                                        else
                                        {
                                            train.mainStation.stoppedTime = "由" + tmp_Train.firstTrainNum + "改";
                                        }
                                    }
                                    else
                                    {//上行，则正好相反
                                        Char[] TrainWord = tmp_Train.firstTrainNum.ToCharArray();
                                        if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                                        {
                                            train.mainStation.stoppedTime = "由" + tmp_Train.firstTrainNum + "改";
                                        }
                                        else
                                        {
                                            train.mainStation.stoppedTime = "由" + tmp_Train.secondTrainNum + "改";
                                        }
                                    }
                                }
                                else if(tmp_Train.firstTrainNum!= null && tmp_Train.firstTrainNum.Length != 0)
                                {
                                    train.mainStation.stoppedTime = "由" + tmp_Train.firstTrainNum + "改";
                                }

                            }
                        }
                        else if (train.mainStation.stationType == 2)
                        {//找始发接续->续开X
                            Train tmp_Train = new Train();
                            foreach (Train compared_Train in trainsWithMainStation)
                            {
                                if (tmp_Train.mainStation == null || tmp_Train.secondTrainNum.Length != 0)
                                {//首次寻找
                                    if (compared_Train.mainStation.stationTrackNum != null && compared_Train.mainStation.stationTrackNum.Length!= 0)
                                    {
                                        if (compared_Train.mainStation.stationTrackNum.Equals(train.mainStation.stationTrackNum))
                                        {//找到了同股道列车-此车必须是始发车
                                            if (compared_Train.mainStation.stationType != 1)
                                            {
                                                continue;
                                            }
                                            int trainTime = 0;
                                            int compared_TrainTime = 0;
                                            //比较-需要找的车的到达时间，和其他车的发车时间，而且发车时间需要大于到达时间
                                            int.TryParse(train.mainStation.stoppedTime.Replace(":", "").Trim(), out trainTime);
                                            int.TryParse(compared_Train.mainStation.startedTime, out compared_TrainTime);

                                            if (compared_TrainTime > trainTime)
                                            {//新找到的时间和原时间差小于之前找到的时间差，用新的替换老的
                                                tmp_Train = compared_Train;
                                            }
                                        }
                                    }
                                }
                                else
                                {//非首次寻找，需要进行比较
                                    if (compared_Train.mainStation.stationTrackNum != null && compared_Train.mainStation.stationTrackNum.Length != 0)
                                    {
                                        if (compared_Train.mainStation.stationTrackNum.Equals(train.mainStation.stationTrackNum))
                                        {//找到了同股道列车-此车必须是始发车
                                            if (compared_Train.mainStation.stationType != 1)
                                            {
                                                continue;
                                            }
                                            int trainTime = 0;
                                            int compared_TrainTime = 0;
                                            int tmp_TrainTime = 0;
                                            //比较-需要找的车的到达时间，和其他车的发车时间，而且发车时间需要大于到达时间
                                            int.TryParse(train.mainStation.stoppedTime.Replace(":", "").Trim(), out trainTime);
                                            int.TryParse(compared_Train.mainStation.startedTime, out compared_TrainTime);
                                            int.TryParse(tmp_Train.mainStation.startedTime, out tmp_TrainTime);
                                            
                                            if ((Math.Abs(trainTime - compared_TrainTime) < Math.Abs(trainTime - tmp_TrainTime)) &&
                                                compared_TrainTime > trainTime)
                                            {//新找到的时间和原时间差小于之前找到的时间差，用新的替换老的
                                                tmp_Train = compared_Train;
                                            }
                                        }
                                    }
                                }
                                //如果有双车次，判断接续前列车的上下行
                                if (tmp_Train.secondTrainNum != null && tmp_Train.secondTrainNum.Length!=0)
                                {
                                    if (tmp_Train.upOrDown)
                                    {//下行
                                        Char[] TrainWord = tmp_Train.firstTrainNum.ToCharArray();
                                        if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                                        {//最后一位是偶数，则用奇数的那个
                                            train.mainStation.startedTime = "续开" + tmp_Train.secondTrainNum;
                                        }
                                        else
                                        {
                                            train.mainStation.startedTime = "续开" + tmp_Train.firstTrainNum;
                                        }
                                    }
                                    else
                                    {//上行，则正好相反
                                        Char[] TrainWord = tmp_Train.firstTrainNum.ToCharArray();
                                        if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                                        {
                                            train.mainStation.startedTime = "续开" + tmp_Train.firstTrainNum;
                                        }
                                        else
                                        {
                                            train.mainStation.startedTime = "续开" + tmp_Train.secondTrainNum;
                                        }
                                    }
                                }
                                else
                                {
                                    train.mainStation.startedTime = "续开" + tmp_Train.firstTrainNum;
                                }
                            }
                        }
                
            }
            return trains;
        }


        //20210602-为每个分表找到经过该车站的列车
        private List<TimeTable> getDistributedTrainsWithALLTRAINS(List<TimeTable> _distributedTimeTables)
        {
            //非疏解区，动车所列车
            //疏解区用经过疏解区-圃田西的时间顺序判定上下行，圃田西->疏解区的在左边，疏解区->圃田西的在右边
            //东所南所单独写，判断走行线
            List<Train> _allTrains = allTrains_New;

            //给主站加上场，再根据各时刻表替换成各时刻表的站

            _allTrains = remarkStations(_allTrains);

            for (int i = 0; i < _distributedTimeTables.Count; i++)
            {
                string mainStation = _distributedTimeTables[i].Title;
                //↓√↑×
                List<Train> _downTrains = new List<Train>();
                List<Train> _upTrains = new List<Train>();
                //其他普通车站，以标题为准，经过该车站的列车都加进去
                //特殊情况：京广场包含马头岗->徐兰场，徐兰场包括东所->城际场，（有了添加）
                for(int  j = 0; j< _allTrains.Count; j++)
                {
                    for(int _sCount = 0; _sCount < _allTrains[j].newStations.Count; _sCount++)
                    {
                        Station _station = _allTrains[j].Clone().newStations[_sCount];
                        Train _tempTrain = _allTrains[j].Clone();
                        if (_station.stationName.Replace("线路所", "").Replace("站", "").Equals(mainStation) || _tempTrain.mainStation.stationName.Equals(mainStation))
                        {
                            //把主站和当前站换一下(如果主站是当前站就不换)
                            if (!_tempTrain.mainStation.stationName.Equals(mainStation))
                            {
                                Station _tempMainStation = _tempTrain.Clone().mainStation;
                                //可能已经在别的时刻表里换过主站了，若newstations里包含，则不添加
                                bool hasGot = false;
                                foreach(Station _s in _tempTrain.newStations)
                                {
                                    if (_s.stationName.Equals(mainStation))
                                    {
                                        hasGot = true;
                                        break;
                                    }
                                }
                                if (!hasGot)
                                {
                                    _tempTrain.newStations.Add(_tempMainStation);
                                }
                                _tempTrain.mainStation = _station;
                            }
                            //特殊情况：动车所，包含“郑州东京广场，郑州东徐兰场，郑州东城际场”的为郑州东站，通过股道判断走行线，包含“郑州南郑万场，郑州南郑阜场，郑州南城际场”的为郑州南站
                            if(_station.stationName.Replace("线路所", "").Replace("站", "").Contains("东动车所"))
                            {
                                //找出郑州东站
                                Station _zzd = new Station();
                                int eastEMUGarageTrackLine = 0;
                                bool hasGot = false;
                                for (int _zzdPlace = 0; _zzdPlace < _tempTrain.newStations.Count; _zzdPlace++)
                                {
                                    if (_tempTrain.newStations[_zzdPlace].stationName.Contains("郑州东京广场")||
                                        _tempTrain.newStations[_zzdPlace].stationName.Contains("郑州东城际场") ||
                                        _tempTrain.newStations[_zzdPlace].stationName.Contains("郑州东徐兰场"))
                                    {
                                        //把原来的复制一份
                                        //郑州南站和郑州南某某场都填进去
                                        Station _clonedStation = new Station();
                                        string _stopTime = _tempTrain.newStations[_zzdPlace].stoppedTime;
                                        string _startTime = _tempTrain.newStations[_zzdPlace].startedTime;
                                        string _track = _tempTrain.newStations[_zzdPlace].stationTrackNum;
                                        string _name = _tempTrain.newStations[_zzdPlace].stationName;
                                        _clonedStation.stoppedTime = _stopTime;
                                        _clonedStation.startedTime = _startTime;
                                        _clonedStation.stationTrackNum = _track;
                                        _clonedStation.stationName = _name;
                                        _tempTrain.newStations[_zzdPlace].stationName = "郑州东站";
                                        _zzd = _tempTrain.newStations[_zzdPlace];
                                        eastEMUGarageTrackLine = getEastEMUGarageTrackLine(_zzd);
                                        _tempTrain.newStations.Add(_clonedStation);
                                        hasGot = true;
                                        break;
                                    }
                                }
                                if (hasGot)
                                {
                                    _station.stationTrackNum = eastEMUGarageTrackLine.ToString();
                                }
                            }
                            //南所，启用完后在此补全
                            if (_station.stationName.Replace("线路所", "").Replace("站", "").Contains("南动车所"))
                            {
                                //找出郑州南站
                                Station _zzn = new Station();
                                int eastEMUGarageTrackLine = 0;
                                bool hasGot = false;
                                for (int _zznPlace = 0; _zznPlace < _tempTrain.newStations.Count; _zznPlace++)
                                {
                                    if (_tempTrain.newStations[_zznPlace].stationName.Contains("郑州南郑万场") ||
                                        _tempTrain.newStations[_zznPlace].stationName.Contains("郑州南郑阜场") ||
                                        _tempTrain.newStations[_zznPlace].stationName.Contains("郑州南城际场"))
                                    {
                                        //郑州南站和郑州南某某场都填进去
                                        Station _clonedStation = new Station();
                                        string _stopTime = _tempTrain.newStations[_zznPlace].stoppedTime;
                                        string _startTime = _tempTrain.newStations[_zznPlace].startedTime;
                                        string _track = _tempTrain.newStations[_zznPlace].stationTrackNum;
                                        string _name = _tempTrain.newStations[_zznPlace].stationName;
                                        _clonedStation.stoppedTime = _stopTime;
                                        _clonedStation.startedTime = _startTime;
                                        _clonedStation.stationTrackNum = _track;
                                        _clonedStation.stationName = _name;

                                        _tempTrain.newStations[_zznPlace].stationName = "郑州南站";
                                        _tempTrain.newStations.Add(_clonedStation);
                                        _zzn = _tempTrain.newStations[_zznPlace];
                                        //eastEMUGarageTrackLine = getEastEMUGarageTrackLine(_zzn);
                                        hasGot = true;
                                        break;
                                    }
                                }
                                if (hasGot)
                                {
                                    //_station.stationTrackNum = eastEMUGarageTrackLine.ToString();
                                }
                            }
                            //特殊情况：如果是疏解区，圃田西->疏解区的在左边（上行），疏解区->圃田西的在右边（下行）
                            //疏解区不能用一般情况添加上下行△
                            bool _isSJQ = false;
                            if (_station.stationName.Replace("线路所", "").Replace("站", "").Contains("疏解区"))
                            {//找一下圃田西站
                                _isSJQ = true;
                                Station _puTianXi = new Station();
                                bool hasGot = false;
                                for(int _ptxPlace = 0; _ptxPlace < _tempTrain.newStations.Count; _ptxPlace++)
                                {
                                    if (_tempTrain.newStations[_ptxPlace].stationName.Contains("圃田西"))
                                    {
                                        _puTianXi = _tempTrain.newStations[_ptxPlace];
                                        hasGot = true;
                                        break;
                                    }
                                }
                                if (hasGot)
                                {//比较方向
                                    int _ptxStartTime = 0;
                                    int.TryParse(_puTianXi.startedTime.Replace(":", "").Trim(), out _ptxStartTime);
                                    int _sjqStartTime = 0;
                                    int.TryParse(_station.startedTime.Replace(":", "").Trim(), out _sjqStartTime);
                                    if (_ptxStartTime < 100 && _ptxStartTime != 0)
                                    {
                                        _ptxStartTime = _ptxStartTime + 2400;
                                    }
                                    if (_sjqStartTime < _ptxStartTime && _sjqStartTime != 0 && _ptxStartTime != 0)
                                    {//下行，放右边
                                        _tempTrain.upOrDown = true;
                                    }
                                    else if (_sjqStartTime > _ptxStartTime && _sjqStartTime != 0 && _ptxStartTime != 0)
                                    {//上行，放左边
                                        _tempTrain.upOrDown = false;
                                    }
                                }
                                if (_tempTrain.upOrDown)
                                {
                                    _downTrains.Add(_tempTrain);
                                }
                                else
                                {
                                    _upTrains.Add(_tempTrain);
                                }
                            }
                            //一般情况，添加上下行
                            if (!_isSJQ)
                            {
                                if (_allTrains[j].upOrDown)
                                {
                                    _downTrains.Add(_tempTrain);
                                }
                                else
                                {
                                    _upTrains.Add(_tempTrain);
                                }
                            }
                        }
                        //特殊情况：马头岗徐兰场
                        if(_station.stationName.Replace("线路所", "").Replace("站", "").Contains("马头岗") && mainStation.Contains("京广场"))
                        {//如果这个车不经过京广场，添加进去
                            bool hasGet = false;
                            foreach(Station _s in _allTrains[j].newStations)
                            {
                                if (_s.stationName.Contains("京广场"))
                                {
                                    hasGet = true;
                                    break;
                                }
                            }
                            if(hasGet == false)
                            {
                                if (_allTrains[j].upOrDown)
                                {
                                    _downTrains.Add(_tempTrain);
                                }
                                else
                                {
                                    _upTrains.Add(_tempTrain);
                                }
                            }
                        }
                        //特殊情况：鸿宝京广场
                        if (_station.stationName.Replace("线路所", "").Replace("站", "").Contains("鸿宝") && mainStation.Contains("徐兰场"))
                        {//如果这个车不经过徐兰场，添加进去
                            bool hasGet = false;
                            foreach (Station _s in _allTrains[j].newStations)
                            {
                                if (_s.stationName.Contains("徐兰场"))
                                {
                                    hasGet = true;
                                    break;
                                }
                            }
                            if (hasGet == false)
                            {
                                if (_allTrains[j].upOrDown)
                                {
                                    _downTrains.Add(_tempTrain);
                                }
                                else
                                {
                                    _upTrains.Add(_tempTrain);
                                }
                            }
                        }
                        //特殊情况：曹古寺京广场
                        if (_station.stationName.Replace("线路所", "").Replace("站", "").Contains("曹古寺") && mainStation.Contains("徐兰场"))
                        {//如果这个车不经过徐兰场，添加进去
                            bool hasGet = false;
                            foreach (Station _s in _allTrains[j].newStations)
                            {
                                if (_s.stationName.Contains("徐兰场"))
                                {
                                    hasGet = true;
                                    break;
                                }
                            }
                            if (hasGet == false)
                            {
                                if (_allTrains[j].upOrDown)
                                {
                                    _downTrains.Add(_tempTrain);
                                }
                                else
                                {
                                    _upTrains.Add(_tempTrain);
                                }
                            }
                        }

                        //特殊情况：东动车所城际场
                        if (_station.stationName.Equals("郑州东动车所") && mainStation.Contains("徐兰场"))
                        {
                            if(_station.stationTrackNum.Equals("17") ||
                                _station.stationTrackNum.Equals("18") ||
                                _station.stationTrackNum.Equals("19") ||
                                _station.stationTrackNum.Equals("20") ||
                                _station.stationTrackNum.Equals("XVIII") ||
                                _station.stationTrackNum.Equals("XIX"))
                            {
                                if (_allTrains[j].upOrDown)
                                {
                                    _downTrains.Add(_tempTrain);
                                }
                                else
                                {
                                    _upTrains.Add(_tempTrain);
                                }
                            }

                        }
                    }
                }
                _distributedTimeTables[i].upTrains = _upTrains;
                _distributedTimeTables[i].downTrains = _downTrains;
                _distributedTimeTables[i].upTrains.Sort();
                _distributedTimeTables[i].downTrains.Sort();

            }
            return _distributedTimeTables;

        }

        //根据股道判断东所走行线
        private int getEastEMUGarageTrackLine(Station _station)
        {
            int mainTrackNum = 0;
            int trackLine = 0;
                if (_station.stationTrackNum.Equals("IX"))
                    mainTrackNum = 9;
                if (_station.stationTrackNum.Equals("X"))
                    mainTrackNum = 10;
                if (_station.stationTrackNum.Equals("XVIII"))
                    mainTrackNum = 18;
                if (_station.stationTrackNum.Equals("XIX"))
                    mainTrackNum = 19;
                if (_station.stationTrackNum.Equals("XXV"))
                    mainTrackNum = 25;
                if (_station.stationTrackNum.Equals("XXVI"))
                    mainTrackNum = 26;
                if (_station.stationTrackNum.Equals("XXIX"))
                    mainTrackNum = 29;
                if (_station.stationTrackNum.Equals("XXX"))
                    mainTrackNum = 30;
                if (mainTrackNum == 0)
                    int.TryParse(_station.stationTrackNum, out mainTrackNum);
                if (mainTrackNum != 0)
                {
                    if (mainTrackNum >= 1 && mainTrackNum <= 10)
                    {
                    trackLine = 1;
                    }
                else if(mainTrackNum >= 11 && mainTrackNum <= 16)
                    {
                    trackLine = 2;
                    }
                    else if (mainTrackNum >= 17 && mainTrackNum <= 25)
                    {
                    trackLine = 3;
                    }
                    else if(mainTrackNum >= 26 && mainTrackNum <= 32)
                {
                    trackLine = 4;
                }
                }
            return trackLine;
        }

        private List<Train> remarkStations(List<Train> _allTrains)
        {
            for (int j = 0; j < _allTrains.Count; j++)
            {
                for (int _sCount = 0; _sCount < _allTrains[j].newStations.Count; _sCount++)
                {
                    int mainTrackNum = 0;
                    if (_allTrains[j].mainStation != null && (_allTrains[j].mainStation.stationName.Equals("郑州东") ||
                        _allTrains[j].mainStation.stationName.Equals("郑州东站")))
                    {
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("IX"))
                            mainTrackNum = 9;
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("X"))
                            mainTrackNum = 10;
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("XVIII"))
                            mainTrackNum = 18;
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("XIX"))
                            mainTrackNum = 19;
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("XXV"))
                            mainTrackNum = 25;
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("XXVI"))
                            mainTrackNum = 26;
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("XXIX"))
                            mainTrackNum = 29;
                        if (_allTrains[j].mainStation.stationTrackNum.Equals("XXX"))
                            mainTrackNum = 30;
                        if (mainTrackNum == 0)
                            int.TryParse(_allTrains[j].mainStation.stationTrackNum, out mainTrackNum);
                        if (mainTrackNum != 0)
                        {
                            if (mainTrackNum >= 1 && mainTrackNum <= 16)
                            {
                                _allTrains[j].mainStation.stationName = "郑州东京广场";
                            }
                            if (mainTrackNum >= 17 && mainTrackNum <= 20)
                            {
                                _allTrains[j].mainStation.stationName = "郑州东城际场";
                            }
                            if (mainTrackNum >= 21 && mainTrackNum <= 30)
                            {
                                _allTrains[j].mainStation.stationName = "郑州东徐兰场";
                            }
                        }
                    }
                }
            }
            return _allTrains;
        }

        //根据表头创建时刻表文件
        //包含-“title”和“时刻表”的是某张表的表名（京广&&时刻表），所在行记录下来
        //包含“始发站”的是车站所在行，行和列都需要记录下来
        //判断这张时刻表是哪张，然后开始从标有“始发”字样的位置向下移动
        //循环寻找时刻表内的车站，发现一个“始发”之后就开始添加
        //此时注意有两个标有“始发”的位置，分别是上下行，
        //先创建行，行数为上下行中车次最多的数量+10
        //直到找到空格子，或者没有字的格子，再判断这个格子右边一格是不是空的，是的话此时开始添加车（一次添加完一个时刻表的上行/下行）
        //根据“始发”的上下行，判断此时添加的是上行还是下行
        //遍历所有该时刻表该方向的车，进行逐一添加
        //添加时，判断这一行有没有数据，没有的话新建一行
        //然后开始填写，填写时，根据车来找时刻表
        //（所有的规则都要匹配上下行）
        //先找车次所在的列，行是新建行，填写车次，然后找始发-终到，填写
        //然后找主站-“京广”或者“徐兰”，在时刻表内寻找相应主站位置，填写
        //然后找其他车站，从其他车站里一个个挑出，填写到该行的对应列上，
        //填写完之后，该行的始发站-终点站之间的格子未被填写的加上斜杠

        private void createTimeTableFile(List<IWorkbook> _timeTableWorkbooks,List<TimeTable> timeTables,int _inputType)
        {
            int timeTablePlace = 0;
            foreach (IWorkbook workbook in _timeTableWorkbooks)
            {
                //用来给表格填斜杠用的
                int _stopColumn = 0;
                //格式-标准
                ICellStyle standard = workbook.CreateCellStyle();
                standard.FillForegroundColor = HSSFColor.White.Index;
                standard.FillPattern = FillPattern.SolidForeground;
                standard.FillBackgroundColor = HSSFColor.White.Index;
                standard.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                standard.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                standard.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                standard.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                standard.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                HSSFFont standardFont = (HSSFFont)workbook.CreateFont();
                standardFont.FontName = "Times New Roman";//字体  
                standardFont.FontHeightInPoints = 15;//字号  
                standard.SetFont(standardFont);

                //格式-续开
                ICellStyle continuedTrainCell = workbook.CreateCellStyle();
                continuedTrainCell.FillForegroundColor = HSSFColor.White.Index;
                continuedTrainCell.FillPattern = FillPattern.SolidForeground;
                continuedTrainCell.FillBackgroundColor = HSSFColor.White.Index;
                continuedTrainCell.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                continuedTrainCell.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                continuedTrainCell.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                continuedTrainCell.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                continuedTrainCell.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                HSSFFont font9B = (HSSFFont)workbook.CreateFont();
                font9B.FontName = "黑体";//字体  
                font9B.FontHeightInPoints = 9;//字号  
                continuedTrainCell.SetFont(font9B);
                /*
                font.Underline = NPOI.SS.UserModel.FontUnderlineType.Double;//下划线  
                font.IsStrikeout = true;//删除线  
                font.IsItalic = true;//斜体  
                font.IsBold = true;//加粗  
                */

                //格式-起点终点
                ICellStyle startAndStop = workbook.CreateCellStyle();
                startAndStop.FillForegroundColor = HSSFColor.White.Index;
                startAndStop.FillPattern = FillPattern.SolidForeground;
                startAndStop.FillBackgroundColor = HSSFColor.White.Index;
                startAndStop.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                startAndStop.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                startAndStop.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                startAndStop.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                startAndStop.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                HSSFFont startAndStopFont = (HSSFFont)workbook.CreateFont();
                startAndStopFont.FontName = "宋体";//字体  
                startAndStopFont.FontHeightInPoints = 13;//字号  
                startAndStop.SetFont(startAndStopFont);

                //格式-车次
                ICellStyle trainNumberCell = workbook.CreateCellStyle();
                trainNumberCell.FillForegroundColor = HSSFColor.White.Index;
                trainNumberCell.FillPattern = FillPattern.SolidForeground;
                trainNumberCell.FillBackgroundColor = HSSFColor.White.Index;
                trainNumberCell.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                trainNumberCell.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                trainNumberCell.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                trainNumberCell.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                trainNumberCell.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                HSSFFont trainNumberFont = (HSSFFont)workbook.CreateFont();
                trainNumberFont.FontName = "Times New Roman";//字体  
                trainNumberFont.FontHeightInPoints = 14;//字号  
                trainNumberCell.SetFont(trainNumberFont);

                //斜杠格式
                ICellStyle empty = workbook.CreateCellStyle();
                empty.BorderDiagonalLineStyle = NPOI.SS.UserModel.BorderStyle.Thin;
                empty.BorderDiagonal = BorderDiagonal.Forward;
                empty.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                empty.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                empty.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                empty.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                empty.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                empty.TopBorderColor = HSSFColor.Black.Index;

                ISheet sheet = workbook.GetSheetAt(0);  //获取工作表  
                //获取和工作表对应的时刻表
                TimeTable table = new TimeTable();
                foreach(TimeTable tb in timeTables)
                {
                    if(tb.timeTablePlace == timeTablePlace)
                    {
                        table = tb;
                    }
                }
                if (table == null)
                {
                    MessageBox.Show("发生运行错误：选择的时刻表表头数量过多");
                    return;
                }
                else
                {
                    bool isEMUGarage = false;
                    if (table.Title.Contains("动车所"))
                    {
                        isEMUGarage = true;
                    }
                    //如果总行数小于车次最多的一列+5的话，创建行直到那么多为止，再顺便创建所有的cell
                    int lastRowNumber = 5;
                    if (table.upTrains.Count > table.downTrains.Count)
                    {
                        lastRowNumber = table.upTrains.Count + 5;
                    }
                    else
                    {
                        lastRowNumber = table.downTrains.Count + 5;
                    }
                    if (sheet.LastRowNum < lastRowNumber)
                    {
                        for (int b = sheet.LastRowNum  + 1; b < lastRowNumber; b++)
                        {
                            sheet.CreateRow(b);
                        }
                    }
                    int LastCellNumber = 0;
                    if (sheet.GetRow(0) != null)
                    {
                        LastCellNumber = sheet.GetRow(0).LastCellNum;
                    }
                    //开始分上下行找
                    for (int i = 0; i < table.currentStations.Count; i++)
                    {
                        //这些row都是可以等于的
                        int titleColumn = 0;
                        int stopColumn = 0;
                        IRow stationRow;
                        int stationRowNumber = 0;
                        bool upOrDown = false;
                        stationRow = sheet.GetRow(table.stationRow);
                        stationRowNumber = table.stationRow;
                        List<Stations_TimeTable> temp_TimeTableStations = table.currentStations;
                        //202106：如果是动车所时刻表
                            //备注是第一列
                        {
                            if ((!isEMUGarage && temp_TimeTableStations [i].stationName.Contains("始发")) ||
                                (isEMUGarage && temp_TimeTableStations[i].stationName.Contains("备注")))
                            {//找到一个始发，运行一次
                                titleColumn = temp_TimeTableStations[i].stationColumn;
                                upOrDown = temp_TimeTableStations[i].upOrDown;
                                bool hasGotStopColumn = false;
                                for (int j = i; j < temp_TimeTableStations.Count; j++)
                                {//找终到
                                    if (temp_TimeTableStations[j].stationName.Contains("终到"))
                                    {
                                        stopColumn = temp_TimeTableStations[j].stationColumn;
                                        _stopColumn = stopColumn;
                                        hasGotStopColumn = true;
                                    }
                                    if (hasGotStopColumn)
                                    {
                                        break;
                                    }
                                }

                                //开始填写
                                int counter = 0;
                                List<Train> _trains = new List<Train>();
                                if (upOrDown)
                                {//下行
                                    _trains = table.downTrains;
                                }
                                else
                                {//上行
                                    _trains = table.upTrains;
                                }
                                //去重-2给1
                                Train _firstTrain = new Train();
                                Train _secondTrain = new Train();
                                for (int t = 0; t < _trains.Count; t++)
                                {
                                    _firstTrain = _trains[t];
                                    //如果是郑州-新郑机场的车，则不参与去重
                                    if ((_firstTrain.startStation.Equals("新郑机场") ||
                                        _firstTrain.startStation.Equals("郑州") ||
                                        _firstTrain.startStation.Contains("焦作")) &&
                                        (_firstTrain.stopStation.Equals("新郑机场") ||
                                        _firstTrain.stopStation.Equals("郑州") ||
                                        _firstTrain.stopStation.Contains("焦作")))
                                    {
                                        continue;
                                    }
                                    for (int tt = t + 1; tt < _trains.Count; tt++)
                                    {
                                        _secondTrain = _trains[tt];
                                        if (_firstTrain.secondTrainNum.Length != 0 && _secondTrain.secondTrainNum.Length != 0)
                                        {
                                            if (_firstTrain.firstTrainNum.Equals(_secondTrain.firstTrainNum) ||
                                                _firstTrain.firstTrainNum.Equals(_secondTrain.secondTrainNum) ||
                                                _firstTrain.secondTrainNum.Equals(_secondTrain.firstTrainNum) ||
                                                _firstTrain.secondTrainNum.Equals(_secondTrain.secondTrainNum))
                                            {//如果车号相同的话，2的元素给1，把2删除
                                                int _sCount = _secondTrain.newStations.Count;
                                                for (int s = 0; s < _sCount; s++)
                                                {
                                                    Station _s = new Station();
                                                    _s.stationName = _secondTrain.newStations[s].stationName;
                                                    _s.startedTime = _secondTrain.newStations[s].startedTime;
                                                    _s.stoppedTime = _secondTrain.newStations[s].stoppedTime;
                                                    _s.stationType = _secondTrain.newStations[s].stationType;
                                                    _s.stationTrackNum = _secondTrain.newStations[s].stationTrackNum;
                                                    Station _ss = new Station();
                                                    _ss = _s;
                                                    _firstTrain.newStations.Add(_ss);
                                                }
                                                _trains.RemoveAt(tt);
                                                tt--;
                                            }
                                        }
                                        else if (_firstTrain.secondTrainNum.Length == 0 && _secondTrain.secondTrainNum.Length == 0)
                                        {//如果两个车相同，要么都有两个车号，要么都有一个车号
                                            if (_firstTrain.firstTrainNum.Equals(_secondTrain.firstTrainNum))
                                            {
                                                /*
                                                for(int ij = 0; ij < _secondTrain.newStations.Count; ij++)
                                                {
                                                    Station _s = _secondTrain.newStations[ij];
                                                    Station _ss = new Station();
                                                    _ss = _s;
                                                    _firstTrain.newStations.Add(_ss);
                                                }
                                                */
                                                foreach (Station _s in _secondTrain.newStations)
                                                {
                                                    Station _ss = new Station();
                                                    _ss = _s;
                                                    bool hasEqual = false;
                                                    foreach (Station _s1 in _firstTrain.newStations)
                                                    {
                                                        if (_s1.stationName.Trim().Equals(_s.stationName))
                                                        {
                                                            hasEqual = true;
                                                        }
                                                    }
                                                    if (!hasEqual)
                                                    {
                                                        _firstTrain.newStations.Add(_ss);
                                                    }
                                                }
                                                _trains.RemoveAt(tt);
                                                tt--;
                                            }
                                        }
                                    }
                                }

                                for (int j = stationRowNumber + 2; j <= sheet.LastRowNum; j++)
                                {//往下找，直接跳过一行
                                    IRow newRow;
                                    if (sheet.GetRow(j) == null)
                                    {
                                        newRow = sheet.CreateRow(j);
                                        for (int m = 0; m < LastCellNumber; m++)
                                        {
                                            newRow.CreateCell(m);
                                        }
                                    }
                                    newRow = sheet.GetRow(j);
                                    if (counter >= _trains.Count)
                                    {
                                        break;
                                    }
                                    Train _train = _trains[counter];
                                    //获取当前上下行时应该填写的车次，疏解区单独写
                                    string trainNumber = "";
                                    if (table.Title.Contains("疏解区"))
                                    {//分为四种，->郑州东（双号），->二郎庙（单号），<-郑州东（单号），<-二郎庙（双号）
                                        trainNumber = _train.firstTrainNum;
                                        bool sjqUpOrDown = false;
                                        if (upOrDown)
                                        {
                                            //下行（疏解区->圃田西）,<-郑州东（单号），<-二郎庙（双号）
                                            for (int _stations = 0; _stations < _train.newStations.Count; _stations++)
                                            {
                                                if (_train.newStations[_stations].stationName.Contains("郑州东"))
                                                {
                                                    sjqUpOrDown = true;
                                                    break;
                                                }
                                                if (_train.newStations[_stations].stationName.Contains("二郎庙"))
                                                {
                                                    sjqUpOrDown = false;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {//上行（圃田西->疏解区），->郑州东（双号），->二郎庙（单号）
                                            for (int _stations = 0; _stations < _train.newStations.Count; _stations++)
                                            {
                                                if (_train.newStations[_stations].stationName.Contains("郑州东"))
                                                {
                                                    sjqUpOrDown = false;
                                                    break;
                                                }
                                                if (_train.newStations[_stations].stationName.Contains("二郎庙"))
                                                {
                                                    sjqUpOrDown = true;
                                                    break;
                                                }
                                            }
                                        }
                                        if (_train.secondTrainNum != null && _train.secondTrainNum.Length != 0)
                                        {
                                            if (_train.secondTrainNum.Length != 0)
                                            {
                                                if (sjqUpOrDown)
                                                {
                                                    Char[] TrainWord = _train.firstTrainNum.ToCharArray();
                                                    if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                                                    {//最后一位是偶数，则用奇数的那个
                                                        trainNumber = _train.secondTrainNum;
                                                    }
                                                    else
                                                    {
                                                        trainNumber = _train.firstTrainNum;
                                                    }
                                                }
                                                else
                                                {
                                                    Char[] TrainWord = _train.firstTrainNum.ToCharArray();
                                                    if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                                                    {
                                                        trainNumber = _train.firstTrainNum;
                                                    }
                                                    else
                                                    {
                                                        trainNumber = _train.secondTrainNum;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    //其他的，用普通查找方法
                                    else
                                    {
                                        trainNumber = returnCorrectUpOrDownNumber(_train, upOrDown);
                                    }

                                    int targetColumn = 0;
                                    //车站的三个数据
                                    int stoppedColumn = 0;
                                    int startedColumn = 0;
                                    int trackNumColumn = 0;
                                    //如果这张时刻表里面没有一个车站能对应上这趟车的任何一个车站，则不能显示在这张表内
                                    /*
                                    if (_inputType == 1)
                                    {
                                        bool skip = true;
                                        foreach (Station _temps in _train.newStations)
                                        {
                                            if (_temps.stationName.Length == 0)
                                            {
                                                continue;
                                            }
                                            for (int m = 0; m < table.stations.Length; m++)
                                            {
                                                if (table.stations[m].Length == 0)
                                                {
                                                    continue;
                                                }
                                                if (_temps.stationName.Contains(table.stations[m]) ||
                                                    (table.stations[m]).Contains(_temps.stationName))
                                                {
                                                    skip = false;
                                                    break;
                                                }
                                            }
                                            if (skip == false)
                                            {
                                                break;
                                            }
                                        }
                                        if (skip)
                                        {
                                            break;
                                        }
                                    }
                                    */

                                    Stations_TimeTable currentStation = new Stations_TimeTable();
                                    if (findColumn(temp_TimeTableStations, "车次", i) != null)
                                    {
                                        currentStation = findColumn(temp_TimeTableStations, "车次", i);
                                        targetColumn = currentStation.stationColumn;
                                    }
                                    //填写车次
                                    if (newRow.GetCell(targetColumn) == null)
                                    {
                                        newRow.CreateCell(targetColumn);
                                    }
                                    newRow.GetCell(targetColumn).CellStyle = trainNumberCell;
                                    newRow.GetCell(targetColumn).SetCellValue(trainNumber);
                                    //找起点
                                    //动车所不写
                                    if (!isEMUGarage)
                                    {

                                        if (findColumn(temp_TimeTableStations, "始发", i) != null)
                                        {
                                            currentStation = findColumn(temp_TimeTableStations, "始发", i);
                                            targetColumn = currentStation.stationColumn;
                                        }
                                        //填写起点
                                        if (newRow.GetCell(targetColumn) == null)
                                        {
                                            newRow.CreateCell(targetColumn);
                                        }
                                        newRow.GetCell(targetColumn).CellStyle = startAndStop;
                                        newRow.GetCell(targetColumn).SetCellValue(_train.startStation);
                                        //找终点
                                        if (findColumn(temp_TimeTableStations, "终到", i) != null)
                                        {
                                            currentStation = findColumn(temp_TimeTableStations, "终到", i);
                                            targetColumn = currentStation.stationColumn;
                                        }
                                        //填写终到站
                                        if (newRow.GetCell(targetColumn) == null)
                                        {
                                            newRow.CreateCell(targetColumn);
                                        }
                                        newRow.GetCell(targetColumn).CellStyle = startAndStop;
                                        newRow.GetCell(targetColumn).SetCellValue(_train.stopStation);
                                    }
                                    else
                                    {//动车所写终到场
                                     //找终点，只有单边有终到场，因此终点站为动车所的不填写即可
                                        if (!_train.stopStation.Contains("动车所"))
                                        {
                                            if (findColumn(temp_TimeTableStations, "终到场", i) != null)
                                            {
                                                currentStation = findColumn(temp_TimeTableStations, "终到场", i);
                                                targetColumn = currentStation.stationColumn;
                                            }
                                            //填写终到场
                                            if (newRow.GetCell(targetColumn) == null)
                                            {
                                                newRow.CreateCell(targetColumn);
                                            }
                                            //把场找回来填到终到站里..原方法使用的是list，懒得改了
                                            //当前找的是东所的，如果是南所另写
                                            int mainTrackNum = 0;
                                            if(_train.newStations.Count != 0 && (_train.stopStation.Equals("郑州东") || _train.stopStation.Equals("郑州东站")))
                                            {
                                                if (_train.newStations[0].stationTrackNum.Equals("IX"))
                                                    mainTrackNum = 9;
                                                if (_train.newStations[0].stationTrackNum.Equals("X"))
                                                    mainTrackNum = 10;
                                                if (_train.newStations[0].stationTrackNum.Equals("XVIII"))
                                                    mainTrackNum = 18;
                                                if (_train.newStations[0].stationTrackNum.Equals("XIX"))
                                                    mainTrackNum = 19;
                                                if (_train.newStations[0].stationTrackNum.Equals("XXV"))
                                                    mainTrackNum = 25;
                                                if (_train.newStations[0].stationTrackNum.Equals("XXVI"))
                                                    mainTrackNum = 26;
                                                if (_train.newStations[0].stationTrackNum.Equals("XXIX"))
                                                    mainTrackNum = 29;
                                                if (_train.newStations[0].stationTrackNum.Equals("XXX"))
                                                    mainTrackNum = 30;
                                                if (mainTrackNum == 0)
                                                    int.TryParse(_train.newStations[0].stationTrackNum, out mainTrackNum);
                                                if (mainTrackNum != 0)
                                                {
                                                    if (mainTrackNum >= 1 && mainTrackNum <= 16)
                                                    {
                                                        _train.stopStation = "京广场";
                                                    }
                                                    if (mainTrackNum >= 17 && mainTrackNum <= 20)
                                                    {
                                                        _train.stopStation = "城际场";
                                                    }
                                                    if (mainTrackNum >= 21 && mainTrackNum <= 30)
                                                    {
                                                        _train.stopStation = "徐兰场";
                                                    }
                                                }
                                            }
                                          
                                            newRow.GetCell(targetColumn).CellStyle = startAndStop;
                                            newRow.GetCell(targetColumn).SetCellValue(_train.stopStation);
                                        }
                                    }

                                    //主站
                                    //徐兰场进京广场的车特殊显示
                                    bool skipThisTrain = false;
                                    if (findColumn(temp_TimeTableStations, table.Title, i) != null)
                                    {
                                        currentStation = findColumn(temp_TimeTableStations, table.Title, i);
                                        if (table.Title.Equals("徐兰"))
                                        {
                                            if (_train.mainStation != null && _train.mainStation.stationName.Length != 0)
                                            {
                                                int outPut = 0;
                                                int.TryParse(_train.mainStation.stationTrackNum, out outPut);
                                                if (outPut > 0 && outPut < 17)
                                                {//京广场的车
                                                    Station _station = new Station();
                                                    _station.stationName = "京广场";
                                                    _station.stoppedTime = _train.mainStation.stoppedTime.ToString();
                                                    _station.startedTime = _train.mainStation.startedTime.ToString();
                                                    _train.newStations.Add(_station);
                                                    skipThisTrain = true;

                                                }
                                            }
                                        }
                                        if (currentStation.trackNumColumn != 0)
                                        {
                                            if (_train.mainStation.stationTrackNum != null && _train.mainStation.stationTrackNum.Length != 0)
                                            {
                                                if (newRow.GetCell(currentStation.trackNumColumn) == null)
                                                {
                                                    newRow.CreateCell(currentStation.trackNumColumn);
                                                }
                                                newRow.GetCell(currentStation.trackNumColumn).CellStyle = standard;
                                                if (skipThisTrain)
                                                {
                                                    //合并左中右三个格子，左格子写上“通过”
                                                    //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                                                    sheet.AddMergedRegion(new CellRangeAddress(j, j, currentStation.stoppedTimeColumn, currentStation.startedTimeColumn));
                                                    if (newRow.GetCell(currentStation.stoppedTimeColumn) == null)
                                                    {
                                                        newRow.CreateCell(currentStation.stoppedTimeColumn);
                                                    }
                                                    if (newRow.GetCell(currentStation.startedTimeColumn) == null)
                                                    {
                                                        newRow.CreateCell(currentStation.startedTimeColumn);
                                                    }
                                                    newRow.GetCell(currentStation.stoppedTimeColumn).SetCellValue("通过");
                                                    newRow.GetCell(currentStation.stoppedTimeColumn).CellStyle = standard;
                                                    newRow.GetCell(currentStation.startedTimeColumn).CellStyle = standard;
                                                }
                                                else
                                                {
                                                    newRow.GetCell(currentStation.trackNumColumn).SetCellValue(_train.mainStation.stationTrackNum);
                                                }

                                            }
                                        }
                                        if (currentStation.stoppedTimeColumn != 0 && !skipThisTrain)
                                        {
                                            if (_train.mainStation.stoppedTime != null && _train.mainStation.stoppedTime.Length != 0)
                                            {
                                                string stoppedTime = "";
                                                if (_train.mainStation.stoppedTime.Contains("排序"))
                                                {
                                                    _train.mainStation.startedTime = " ";
                                                    stoppedTime = " ";
                                                }
                                                else
                                                {
                                                    stoppedTime = _train.mainStation.stoppedTime;
                                                }
                                                if (newRow.GetCell(currentStation.stoppedTimeColumn) == null)
                                                {
                                                    newRow.CreateCell(currentStation.stoppedTimeColumn);
                                                }
                                                if (_train.mainStation.stoppedTime.Contains("改"))
                                                {
                                                    newRow.GetCell(currentStation.stoppedTimeColumn).CellStyle = continuedTrainCell;
                                                }
                                                else
                                                {
                                                    newRow.GetCell(currentStation.stoppedTimeColumn).CellStyle = standard;
                                                }

                                                newRow.GetCell(currentStation.stoppedTimeColumn).SetCellValue(stoppedTime);
                                            }
                                        }
                                        //填写主站（默认情况）
                                        if (currentStation.startedTimeColumn != 0 && !skipThisTrain)
                                        {
                                            if (_train.mainStation.startedTime != null && _train.mainStation.startedTime.Length != 0)
                                            {
                                                if (newRow.GetCell(currentStation.startedTimeColumn) == null)
                                                {
                                                    newRow.CreateCell(currentStation.startedTimeColumn);
                                                }
                                                if (_train.mainStation.startedTime.Contains("续开"))
                                                {
                                                    newRow.GetCell(currentStation.startedTimeColumn).CellStyle = continuedTrainCell;
                                                }
                                                else
                                                {
                                                    newRow.GetCell(currentStation.startedTimeColumn).CellStyle = standard;
                                                }
                                                newRow.GetCell(currentStation.startedTimeColumn).SetCellValue(addColonToStartTime(_train.mainStation.startedTime));
                                            }
                                        }


                                    }

                                    //
                                    //
                                    //
                                    //
                                    //填写模块
                                    //
                                    //
                                    //
                                    //
                                    //
                                    //根据每个站进行匹配填写

                                    foreach (Station _station in _train.newStations)
                                    {
                                        if (table.Title.Contains("寺后"))
                                        {
                                            if (_station.stationName.Contains("郑州南"))
                                            {
                                                int aaa = 0;
                                            }

                                        }
                                        if (table.Title.Contains("寺后") && _station.stationName.Contains("郑州南城际场"))
                                        {
                                            int aaa = 0;
                                        }
                                        if (findColumn(temp_TimeTableStations, table.Title, i) != null)
                                        {
                                            currentStation = findColumn(temp_TimeTableStations, _station.stationName, i);
                                            if (currentStation.stoppedTimeColumn != 0)
                                            {
                                                if (_station.stoppedTime != null && _station.stoppedTime.Length != 0)
                                                {
                                                    string stoppedTime = "";
                                                    stoppedTime = _station.stoppedTime;
                                                    if (newRow.GetCell(currentStation.stoppedTimeColumn) == null)
                                                    {
                                                        newRow.CreateCell(currentStation.stoppedTimeColumn);
                                                    }
                                                    if (stoppedTime.Contains("通过") &&
                                                        currentStation.startedTimeColumn == 0)
                                                    {//如果是只显示一个时间(只显示“到达”)，但是又要都显示上时间的话（不能显示“通过”）
                                                        stoppedTime = _station.startedTime;
                                                    }
                                                    newRow.GetCell(currentStation.stoppedTimeColumn).CellStyle = standard;
                                                    newRow.GetCell(currentStation.stoppedTimeColumn).SetCellValue(stoppedTime);
                                                }
                                            }
                                            if (currentStation.startedTimeColumn != 0)
                                            {
                                                if (_station.startedTime != null && _station.startedTime.Length != 0)
                                                {
                                                    string startedTime = _station.startedTime;
                                                    if (newRow.GetCell(currentStation.startedTimeColumn) == null)
                                                    {
                                                        newRow.CreateCell(currentStation.startedTimeColumn);
                                                    }
                                                    if (startedTime.Contains("终到") &&
                                                        currentStation.stoppedTimeColumn == 0)
                                                    {//如果是只显示一个时间(只显示“发出”)，但是又要都显示上时间的话（不能显示“终到”）
                                                        startedTime = _station.stoppedTime;
                                                    }
                                                    newRow.GetCell(currentStation.startedTimeColumn).CellStyle = standard;
                                                    newRow.GetCell(currentStation.startedTimeColumn).SetCellValue(addColonToStartTime(startedTime));
                                                }
                                            }
                                            if (currentStation.trackNumColumn != 0)
                                            {
                                                if (_station.stationTrackNum != null && _station.stationTrackNum.Length != 0)
                                                {
                                                    if (newRow.GetCell(currentStation.trackNumColumn) == null)
                                                    {
                                                        newRow.CreateCell(currentStation.trackNumColumn);
                                                    }
                                                    newRow.GetCell(currentStation.trackNumColumn).CellStyle = standard;
                                                    newRow.GetCell(currentStation.trackNumColumn).SetCellValue(_station.stationTrackNum);
                                                }
                                            }

                                        }
                                    }
                                    counter++;
                                }
                            }
                        }

                        }
                    /*重新修改文件指定单元格样式*/
                    //空的加斜杠，但动车所的表加空格
                        for(int i = 0; i <= sheet.LastRowNum; i++)
                        {
                        if(sheet.GetRow(i) != null)
                        {
                            IRow _row = sheet.GetRow(i);
                            if(_row.GetCell(1) == null)
                            {
                                continue;
                            }else if(_row.GetCell(1).ToString().Trim().Length == 0)
                            {
                                continue;
                            }
                            for (int j = 0; j < _stopColumn; j++)
                            {
                                if(_row.GetCell(j) == null)
                                {
                                    _row.CreateCell(j);
                                }
                                if (!_row.GetCell(j).IsMergedCell)
                                {
                                    if (_row.GetCell(j).ToString().Trim().Length == 0)
                                    {
                                        if (isEMUGarage)
                                        {
                                            _row.GetCell(j).CellStyle = startAndStop;
                                        }
                                        else
                                        {
                                            _row.GetCell(j).CellStyle = empty;
                                        }

                                    }
                                }
                            }
                        }

                        }
                        try
                        {
                            FileStream file = new FileStream(table.fileName.Replace(table.fileName.Split('\\')[table.fileName.Split('\\').Length - 1], "处理后-"+table.fileName.Split('\\')[table.fileName.Split('\\').Length - 1]) , FileMode.Create);
                            //System.Diagnostics.Process.Start("explorer.exe", table.fileName.Replace(table.fileName.Split('\\')[table.fileName.Split('\\').Length - 1], "处理后-" + table.fileName.Split('\\')[table.fileName.Split('\\').Length - 1]));
                            workbook.Write(file);

                            file.Close();
                        }
                        catch (IOException e)
                        {
                            MessageBox.Show("被创建的时刻表正在使用" + "\n错误内容：" + e.ToString().Split('在')[0], "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }

                        timeTablePlace++;
                    }
            }
        }

        private string addColonToStartTime(string startTime)
        {//填写前为主站的开车时间加上冒号
         //先判断是不是纯数字
            Regex r = new Regex(@"^[0-9]+$");
            if (r.Match(startTime).Success)
            {//仅包含数字
                //判断几位
                if (startTime.Length == 1)
                {//一位数字
                    startTime = "0:0" + startTime;
                }
                else if (startTime.Length == 2)
                {
                    startTime = "0:" + startTime;
                }
                else if (startTime.Length == 3)
                {
                    char[] startChar = startTime.ToCharArray();
                    startTime = startChar[0].ToString() + ":" + startChar[1].ToString() + startChar[2].ToString();
                }
                else if (startTime.Length == 4)
                {
                    char[] startChar = startTime.ToCharArray();
                    startTime = startChar[0].ToString() + startChar[1].ToString() + ":" + startChar[2].ToString() + startChar[3].ToString();
                }
            }
                return startTime;
        }

        private Stations_TimeTable findColumn(List<Stations_TimeTable> temp_TimeTableStations,string searchedName, int startColumn)
        {//找列名对应的列的车站
            Stations_TimeTable targetColumn = new Stations_TimeTable();
            bool hasGotTheColumn = false;
            for (int q = startColumn; q < temp_TimeTableStations.Count; q++)
            {//在这里面找带关键字的列
                string stationName = temp_TimeTableStations[q].stationName;
                if (stationName.Contains(searchedName) ||
                    searchedName.Contains(stationName))
                {//这一列就是目标列
                 //郑州南城际场/郑州东城际场区分，动车所区分
                    /*
                    if(searchedName.Trim().Equals("郑州南城际场") && stationName.Equals("城际场"))
                    {
                        continue;
                    }
                    */
                    if (searchedName.Trim().Equals("郑州"))
                    {
                        continue;
                    }
                    if (searchedName.Trim().Equals("郑州东动车所") && stationName.Equals("郑州东"))
                    {
                        continue;
                    }
                    if (searchedName.Trim().Equals("郑州南动车所") && stationName.Equals("郑州南"))
                    {
                        continue;
                    }
                    if (searchedName.Trim().Equals("郑州东疏解区") && stationName.Equals("郑州东"))
                    {
                        continue;
                    }
                    targetColumn = temp_TimeTableStations[q];
                    hasGotTheColumn = true;
                }
                if (hasGotTheColumn)
                {
                    break;
                }
            }
            return targetColumn;
        }

        private string returnCorrectUpOrDownNumber(Train tmp_Train, bool upOrDown)
        {//返回正确的上下行对应的车次
            string trainNumber = "";
            trainNumber = tmp_Train.firstTrainNum;
            if(tmp_Train.secondTrainNum != null&& tmp_Train.secondTrainNum.Length!=0)
            {
                if(tmp_Train.secondTrainNum.Length != 0)
                {
                    if (upOrDown)
                    {//下行
                        Char[] TrainWord = tmp_Train.firstTrainNum.ToCharArray();
                        if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                        {//最后一位是偶数，则用奇数的那个
                            trainNumber = tmp_Train.secondTrainNum;
                        }
                        else
                        {
                            trainNumber = tmp_Train.firstTrainNum;
                        }
                    }
                    else
                    {//上行，则正好相反
                        Char[] TrainWord = tmp_Train.firstTrainNum.ToCharArray();
                        if (TrainWord[TrainWord.Length - 1] % 2 == 0)
                        {
                            trainNumber = tmp_Train.firstTrainNum;
                        }
                        else
                        {
                            trainNumber = tmp_Train.secondTrainNum;
                        }
                    }
                }
            }
            return trainNumber;
        }

        private void ImportNewTimeTable_btn_Click(object sender, EventArgs e)
        {
            ImportFiles(0);
        }

        private void ImportCurrentTimeTable_btn_Click(object sender, EventArgs e)
        {
            ImportFiles(1);
        }

        private void getTrains_btn_Click(object sender, EventArgs e)
        {
            if(NewTimeTablesWorkbooks != null && DistributedTimeTableWorkbooks != null)
            {
                allTrains_New = new List<Train>();
                List<TimeTable> _dt = GetStationsFromCurrentTables(DistributedTimeTableWorkbooks, allDistributedTimeTables, 1);
                if (_dt != null)
                {
                    allDistributedTimeTables = _dt;
                    List<Train> _tempTrains = new List<Train>();
                    //普通临客表
                    if(selectNewTimeTableMode == 0)
                    {
                        _tempTrains = GetTrainsFromNewTimeTables();
                    }
                    //子东表
                    else if(selectNewTimeTableMode == 1)
                    {
                        _tempTrains = ZiDongVersion_GetTrainsFromNewTimeTables();
                    }
                    //四大表
                    else if(selectNewTimeTableMode == 2)
                    {

                    }
                    //原版路局时刻表
                    else if(selectNewTimeTableMode == 3)
                    {

                    }
                    analyizeTrainData(_tempTrains);
                    //把每个分表的车匹配上
                    allTrains_New = _tempTrains;
                    allDistributedTimeTables = getDistributedTrainsWithALLTRAINS(_dt);
                    //matchTrainAndTimeTable();
                    if (_tempTrains != null && _tempTrains.Count != 0)
                    {
                        allTrains_New = _tempTrains;
                        createTimeTableFile(DistributedTimeTableWorkbooks, _dt, 1);
                    }
                    else
                    {
                        MessageBox.Show("未读取到任何车次");
                        return;
                    }
                }
                MessageBox.Show("处理完成","提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("未选择文件");
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }

        //
        //
        //
        //分表填写
        //
        //
        //
        //导入分表文件
        private void ImportDistributedTrainTimeTableFile_btn_Click(object sender, EventArgs e)
        {
            if (ImportFiles(2))
            {
                //导入成功，有分表文件
                hasDistributedTimeTable = true;

            }
        }

        //选择模式
        private void modeSelect_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectNewTimeTableMode = modeSelect_cb.SelectedIndex;
        }


        //分表
    }
}
