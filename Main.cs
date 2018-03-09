using AutomaticTimeTableMakingTools.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
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

namespace AutomaticTimeTableMakingTools
{
    public partial class Main : Form
    {
        List<IWorkbook> NewTimeTablesWorkbooks;
        List<IWorkbook> CurrentTimeTablesWorkbooks;
        List<Train> allTrains_New = new List<Train>();
        string passingStations = "";
        
        public Main()
        {
            InitializeComponent();
            initUI();
        }

        private void initUI()
        {
            newTrains_lv.View = View.Details;
            string[] informationTitle = new string[] { "车次1", "车次2","始发-终到","上/下行" ,"停站信息"};
            this.newTrains_lv.BeginUpdate();
            for (int i = 0; i < 5; i++)
            {
                ColumnHeader ch = new ColumnHeader();
                ch.Text = informationTitle[i];   //设置列标题 
                if(i == 0 || i == 1 || i ==3)
                {
                    ch.Width = 55;
                }
                else if(i == 2)
                {
                    ch.Width = 100;
                }
                else
                {
                    ch.Width = 1000;
                }
                this.newTrains_lv.Columns.Add(ch);    //将列头添加到ListView控件。
            }

            this.newTrains_lv.EndUpdate();
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
                    fileNames = fileNames + "\n" + fileName;
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
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("选中的部分时刻表文件正在使用中，请关闭后重试\n" + fileName, "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return false;
                    }
                }
                switch (fileType)
                {//0为新，1为当前
                    case 0:
                        this.NewTimeTableFile_lbl.Text = fileNames;
                        this.NewTimeTablesWorkbooks = workBooks;
                        break;
                    case 1:
                        this.CurrentTimeTableFile_lbl.Text = fileNames;
                        this.CurrentTimeTablesWorkbooks = workBooks;
                        break;
                }
            }
            return true;
        }

        private void GetStationsFromCurrentTables()
        {
            //在当前的京广徐兰时刻表内找出管辖内所有的车站
            string allStations = "";
            foreach (IWorkbook workbook in CurrentTimeTablesWorkbooks)
            {
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
                                if (row.GetCell(j).ToString().Contains("始发") || hasGotStationsRow)
                                {
                                    hasGotStationsRow = true;
                                    if (!row.GetCell(j).ToString().Contains("始发") &&
                                        !row.GetCell(j).ToString().Contains("终到") &&
                                        !row.GetCell(j).ToString().Contains("车"))
                                    {
                                        string currentStation = row.GetCell(j).ToString();
                                        if (currentStation.Contains("线路所"))
                                            currentStation = currentStation.Replace("线路所", " ");
                                        if (currentStation.Contains("站"))
                                            currentStation = currentStation.Replace("站", " ");
                                        if (currentStation.Contains("郑州东"))
                                            currentStation = currentStation.Replace("郑州东", " ");
                                        currentStation = currentStation.Trim();
                                        if (!allStations.Contains(currentStation))
                                        {
                                            allStations = allStations + "," + currentStation;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (hasGotStationsRow)
                    {
                        break;
                    }
                }
            }
            stations_tb.Text = allStations;
            passingStations = allStations;
        }

        private void GetTrainsFromNewTimeTables()
        {
            /*
            //重新修改文件指定单元格样式
            FileStream fs1 = File.OpenWrite(ExcelFile.FileName);
            workbook.Write(fs1);
            fs1.Close();
            Console.ReadLine();
            fileStream.Close();
            workbook.Close();
            */
            //对于每一个工作簿，先把左边一列的列位置找出，然后在时刻表中根据行来确定车站名称
            //发现“车次”字样后，右边的都是车次，根据车次所在位置往上找，(为空的向左上找)若左边对应列为“终到”，则为终到站，若为“始发”，则为始发站。
            //双车次在检测到车次的时候就进行分离，第一车次需要和上/下行对应（寻找周边的车次）
            //找到每一个车次后，直接对该车次的时刻表/股道进行添加，若往右已经没有了则结束。
            //当找到下一个“始发站”的时候，意味着是下一组车次。
            List<Train> trains = new List<Train>();
            foreach (IWorkbook workbook in NewTimeTablesWorkbooks)
            {
                for(int i = 0; i < workbook.NumberOfSheets; i++)
                {//获取所有工作表
                    ISheet sheet = workbook.GetSheetAt(i);
                    IRow row;
                    //表头数据
                    int[] _startStationRaw = new int[10];
                    int[] _stopStationRaw = new int[10];
                    int trainRawCounter = 0;
                    int[] _trainRawNum = new int[10];
                    int titleColumn = 0;
                    //已经找到，不再继续找
                    //bool shouldContinue = true;
                    //上行双数false 下行单数true
                    bool[] upOrDown = new bool[10];
                    for (int j = 0; j < sheet.LastRowNum; j++)
                    {//找表头数据
                        row = sheet.GetRow(j);
                        if(row != null)
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
                    for (int m = 0; m < _trainRowNumLen; m++)
                    {//找每一个车次行是上行还是下行
                        IRow trainRow = sheet.GetRow(_trainRawNum[m]);
                        //判断有没有找齐两个时间
                        int _continueCounter_hour = 0;
                        int _continueCounter_minute = 0;
                        //两个时间，如果有小时的话优先用小时，否则用分钟对比
                        int[] _trainTimeWithHour = new int[2];
                        int[] _trainTimeWithMinute = new int[2];
                        for (int t = titleColumn + 1; t < trainRow.LastCellNum; t++)
                        {//t为列，firstTrainRaw为行
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
                                    else if(m < _trainRowNumLen - 1)
                                    {
                                        _loadedLastRaw = _startStationRaw[m+1];
                                    }
                                    else
                                    {
                                        return;
                                    }
                                        for (int tt = _trainRawNum[m]; tt < _loadedLastRaw; tt++)
                                        {
                                            IRow tempRaw = sheet.GetRow(tt);
                                            string cellInfo = "";
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
                                                else if (cellInfo.Trim().Length > 0&&
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
                                    }
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
                            else if(_trainTimeWithMinute[0] > _trainTimeWithMinute[1])
                            {//上行
                                upOrDown[m] = false;
                            }
                        }
                        else
                        {
                            if (_trainTimeWithHour[0] < _trainTimeWithHour[1])
                            {//下行
                                upOrDown[m] = true;
                            }
                            else
                            {//上行
                                upOrDown[m] = false;
                            }
                        }
                    }
                  
                   
                    //=========
                    //开始添加该sheet中的车次
                    //=========
                    for(int _rowNum = 0; _rowNum < _trainRowNumLen; _rowNum++)
                    {
                        if(_trainRawNum[_rowNum] == 0)
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
                                            //此处预留找起点站终点站的
                                            //位置
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
                                                if(stopRow.GetCell(s) != null)
                                                {
                                                    stopStation = stopRow.GetCell(s).ToString().Trim();
                                                    if (!stopStation.Equals("")&&
                                                        !stopStation.Contains("终到")&&
                                                        continueFindingStop)
                                                    {
                                                        tempTrain.stopStation = stopStation.Trim();
                                                        continueFindingStop = false;
                                                    }
                                                }
                                                if(startRow.GetCell(s) != null)
                                                {
                                                    startStation = startRow.GetCell(s).ToString().Trim();
                                                    if (!startStation.Equals("") &&
                                                        !startStation.Contains("始发")&&
                                                        continueFindingStart)
                                                    {
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
                                            tempTrain.upOrDown = upOrDown[_rowNum];
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
                                            if(_rowNum == _trainRowNumLen - 1)
                                            {//判断搜索的下边界
                                                stoppedRow = sheet.LastRowNum;
                                            }
                                            else
                                            {
                                                stoppedRow = _startStationRaw[_rowNum + 1];
                                            }
                                            if (upOrDown[_rowNum])
                                            {//下行
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
                                                //判断一下是不是始发站
                                                bool onlyStart = false;
                                                for(int tt = _trainRawNum[_rowNum];tt<= stoppedRow; tt++)
                                                {
                                                    string _foundTime = "";
                                                    IRow _trainTimeRow = sheet.GetRow(tt);
                                                   if( _trainTimeRow.GetCell(t) != null)
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
                                                        else if(cellInfo.Trim().Length > 0)
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
                                                                                //是发点
                                                                                _tempStartingTime = _foundTime;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if(!_foundTime.Equals(_tempStartingTime)&&
                                                           !_foundTime.Equals(_tempStartingTime))
                                                        {//如果时间没有填进去？？
                                                            
                                                        }
                                                        string tempTrack = "";
                                                        if (t < tempRow.LastCellNum - 1)
                                                        {
                                                            if(_trainTimeRow.GetCell(t + 1) != null)
                                                            if (!_trainTimeRow.GetCell(t).ToString().Trim().Equals(""))
                                                            {//找股道
                                                                tempTrack = _trainTimeRow.GetCell(t+1).ToString().Trim();
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
                                                        if (_tempStartingTime.Length != 0&&
                                                            _tempStoppedTime.Length != 0)
                                                        {//已经获取到一组数据，此时应当添加模型
                                                            _tempStation.stoppedTime = _tempStoppedTime;
                                                            _tempStation.startedTime = _tempStartingTime;
                                                            _tempStation.stationType = _stationType;
                                                            _tempStation.stationName = _stationName;
                                                            _tempStation.stationTrackNum = _track;
                                                            tempStations.Add(_tempStation);
                                                            //清零
                                                            _tempStartingTime = "";
                                                            _tempStoppedTime = "";
                                                            _stationType = 0;
                                                            _stationName = "";
                                                            _track = "";
                                                        }
                                                        else if(_tempStartingTime.Length != 0&&
                                                            _tempStoppedTime.Length == 0)
                                                        {//只有发时，没有停时，此站为始发站
                                                            _tempStation.startedTime = _tempStartingTime;
                                                            _tempStation.stoppedTime = "始发";
                                                            _tempStation.stationType = 1;
                                                            _tempStation.stationName = _stationName;
                                                            _tempStation.stationTrackNum = _track;
                                                            tempStations.Add(_tempStation);
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
                                                    if(_trainTimeRow != null)
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
                                                            if(cellInfo.Trim().Length > 0 &&
                                                                    !_hours.Equals("")&&
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

                                                            _tempStation.stoppedTime = _tempStoppedTime;
                                                            _tempStation.startedTime = _tempStartingTime;
                                                            _tempStation.stationType = _stationType;
                                                            _tempStation.stationName = _stationName;
                                                            _tempStation.stationTrackNum = _track;
                                                            tempStations.Add(_tempStation);
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
                                                            _tempStation.startedTime = _tempStartingTime;
                                                            _tempStation.stoppedTime = "始发";
                                                            _tempStation.stationType = 1;
                                                            _tempStation.stationName = _stationName;
                                                            _tempStation.stationTrackNum = _track;
                                                            tempStations.Add(_tempStation);
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
                                            tempTrain.newStations = tempStations;
                                            this.allTrains_New.Add(tempTrain);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            this.newTrains_lv.BeginUpdate();
            foreach (Train model in allTrains_New)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.SubItems[0].Text = model.firstTrainNum;
                lvi.SubItems.Add(model.secondTrainNum);
                lvi.SubItems.Add(model.startStation + "-" + model.stopStation);
                if (model.upOrDown)
                {
                    lvi.SubItems.Add("下行");
                }
                else
                {
                    lvi.SubItems.Add("上行");
                }
                string stations = "";
                foreach(Station mStation in model.newStations)
                {
                    stations = stations + "||" + mStation.stationName + "-" + mStation.stoppedTime + "-" + mStation.startedTime + "\n";
                }
                lvi.SubItems.Add(stations);

                this.newTrains_lv.Items.Add(lvi);
            }
            this.newTrains_lv.EndUpdate();
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
            if(NewTimeTablesWorkbooks != null && CurrentTimeTablesWorkbooks != null)
            {
                GetTrainsFromNewTimeTables();
            }
            else
            {
                MessageBox.Show("未选择文件");
            }
        }
    }
}
