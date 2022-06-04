namespace AutomaticTimeTableMakingTools
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.ImportNewTimeTable_btn = new System.Windows.Forms.Button();
            this.NewTimeTableFile_lbl = new System.Windows.Forms.Label();
            this.getTrains_btn = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.ImportDistributedTrainTimeTableFile_btn = new System.Windows.Forms.Button();
            this.DistributedTimeTableFile_lbl = new System.Windows.Forms.Label();
            this.modeSelect_cb = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.process_lbl = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ImportNewTimeTable_btn
            // 
            this.ImportNewTimeTable_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ImportNewTimeTable_btn.Location = new System.Drawing.Point(64, 166);
            this.ImportNewTimeTable_btn.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ImportNewTimeTable_btn.Name = "ImportNewTimeTable_btn";
            this.ImportNewTimeTable_btn.Size = new System.Drawing.Size(175, 58);
            this.ImportNewTimeTable_btn.TabIndex = 0;
            this.ImportNewTimeTable_btn.Text = "需转换的时刻表";
            this.ImportNewTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportNewTimeTable_btn.Click += new System.EventHandler(this.ImportNewTimeTable_btn_Click);
            // 
            // NewTimeTableFile_lbl
            // 
            this.NewTimeTableFile_lbl.AutoSize = true;
            this.NewTimeTableFile_lbl.Location = new System.Drawing.Point(106, 263);
            this.NewTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.NewTimeTableFile_lbl.Name = "NewTimeTableFile_lbl";
            this.NewTimeTableFile_lbl.Size = new System.Drawing.Size(0, 18);
            this.NewTimeTableFile_lbl.TabIndex = 2;
            // 
            // getTrains_btn
            // 
            this.getTrains_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.getTrains_btn.Location = new System.Drawing.Point(566, 166);
            this.getTrains_btn.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.getTrains_btn.Name = "getTrains_btn";
            this.getTrains_btn.Size = new System.Drawing.Size(175, 58);
            this.getTrains_btn.TabIndex = 4;
            this.getTrains_btn.Text = "开始转换";
            this.getTrains_btn.UseVisualStyleBackColor = true;
            this.getTrains_btn.Click += new System.EventHandler(this.getTrains_btn_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label5.Location = new System.Drawing.Point(202, 344);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(352, 24);
            this.label5.TabIndex = 19;
            this.label5.Text = "时刻表表头仅支持Excel 2003文件（*.xls）";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.OrangeRed;
            this.label6.Location = new System.Drawing.Point(487, 395);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(190, 31);
            this.label6.TabIndex = 20;
            this.label6.Text = "Build 20220604";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label7.Location = new System.Drawing.Point(44, 317);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(680, 24);
            this.label7.TabIndex = 21;
            this.label7.Text = "在各表头添加/删除中间站/线路所，注意表头格式（关键字：到达-股道-发出-通过）";
            // 
            // ImportDistributedTrainTimeTableFile_btn
            // 
            this.ImportDistributedTrainTimeTableFile_btn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ImportDistributedTrainTimeTableFile_btn.Location = new System.Drawing.Point(320, 166);
            this.ImportDistributedTrainTimeTableFile_btn.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ImportDistributedTrainTimeTableFile_btn.Name = "ImportDistributedTrainTimeTableFile_btn";
            this.ImportDistributedTrainTimeTableFile_btn.Size = new System.Drawing.Size(175, 58);
            this.ImportDistributedTrainTimeTableFile_btn.TabIndex = 22;
            this.ImportDistributedTrainTimeTableFile_btn.Text = "各行车岗点空表头";
            this.ImportDistributedTrainTimeTableFile_btn.UseVisualStyleBackColor = true;
            this.ImportDistributedTrainTimeTableFile_btn.Click += new System.EventHandler(this.ImportDistributedTrainTimeTableFile_btn_Click);
            // 
            // DistributedTimeTableFile_lbl
            // 
            this.DistributedTimeTableFile_lbl.AutoSize = true;
            this.DistributedTimeTableFile_lbl.Location = new System.Drawing.Point(334, 263);
            this.DistributedTimeTableFile_lbl.Name = "DistributedTimeTableFile_lbl";
            this.DistributedTimeTableFile_lbl.Size = new System.Drawing.Size(17, 18);
            this.DistributedTimeTableFile_lbl.TabIndex = 23;
            this.DistributedTimeTableFile_lbl.Text = " ";
            // 
            // modeSelect_cb
            // 
            this.modeSelect_cb.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.modeSelect_cb.FormattingEnabled = true;
            this.modeSelect_cb.Items.AddRange(new object[] {
            "①路局表(每趟单独显示)",
            "②技术科新时刻表"});
            this.modeSelect_cb.Location = new System.Drawing.Point(23, 114);
            this.modeSelect_cb.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.modeSelect_cb.Name = "modeSelect_cb";
            this.modeSelect_cb.Size = new System.Drawing.Size(242, 32);
            this.modeSelect_cb.TabIndex = 24;
            this.modeSelect_cb.SelectedIndexChanged += new System.EventHandler(this.modeSelect_cb_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(20, 80);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 24);
            this.label1.TabIndex = 25;
            this.label1.Text = "类型";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label2.Location = new System.Drawing.Point(31, 294);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(701, 24);
            this.label2.TabIndex = 26;
            this.label2.Text = "提示：使用技术科时刻表时，将需要使用的时刻表复制在新Excel文件内并放在首个标签";
            // 
            // process_lbl
            // 
            this.process_lbl.AutoSize = true;
            this.process_lbl.Location = new System.Drawing.Point(588, 248);
            this.process_lbl.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.process_lbl.Name = "process_lbl";
            this.process_lbl.Size = new System.Drawing.Size(0, 18);
            this.process_lbl.TabIndex = 28;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.OrangeRed;
            this.label3.Location = new System.Drawing.Point(409, 426);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(332, 18);
            this.label3.TabIndex = 29;
            this.label3.Text = "修改：航空港站更名，车站名称自动获取";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(777, 469);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.process_lbl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.modeSelect_cb);
            this.Controls.Add(this.DistributedTimeTableFile_lbl);
            this.Controls.Add(this.ImportDistributedTrainTimeTableFile_btn);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.getTrains_btn);
            this.Controls.Add(this.NewTimeTableFile_lbl);
            this.Controls.Add(this.ImportNewTimeTable_btn);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Main";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ImportNewTimeTable_btn;
        private System.Windows.Forms.Label NewTimeTableFile_lbl;
        private System.Windows.Forms.Button getTrains_btn;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button ImportDistributedTrainTimeTableFile_btn;
        private System.Windows.Forms.Label DistributedTimeTableFile_lbl;
        private System.Windows.Forms.ComboBox modeSelect_cb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label process_lbl;
        private System.Windows.Forms.Label label3;
    }
}

