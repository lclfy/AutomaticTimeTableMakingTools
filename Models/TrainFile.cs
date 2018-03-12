using System;
using System.Collections.Generic;
using System.Text;

namespace AutomaticTimeTableMakingTools.Models
{
    public class TrainFile
    {
        public string fileName { get; set; }
        public string workbookNum { get; set; }
        //使用-分割（x行x列）
        public string columnsAndRaws { get; set; }
    }
}
