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
            this.ImportCurrentTimeTable_btn = new System.Windows.Forms.Button();
            this.NewTimeTableFile_lbl = new System.Windows.Forms.Label();
            this.CurrentTimeTableFile_lbl = new System.Windows.Forms.Label();
            this.getTrains_btn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.newTrains_lv = new System.Windows.Forms.ListView();
            this.label3 = new System.Windows.Forms.Label();
            this.currentTimeTableStation_tb = new System.Windows.Forms.RichTextBox();
            this.trainCount_lb = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.ImportDistributedTrainTimeTableFile_btn = new System.Windows.Forms.Button();
            this.DistributedTimeTableFile_lbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ImportNewTimeTable_btn
            // 
            this.ImportNewTimeTable_btn.Location = new System.Drawing.Point(219, 821);
            this.ImportNewTimeTable_btn.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.ImportNewTimeTable_btn.Name = "ImportNewTimeTable_btn";
            this.ImportNewTimeTable_btn.Size = new System.Drawing.Size(194, 81);
            this.ImportNewTimeTable_btn.TabIndex = 0;
            this.ImportNewTimeTable_btn.Text = "所有路局新时刻表";
            this.ImportNewTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportNewTimeTable_btn.Click += new System.EventHandler(this.ImportNewTimeTable_btn_Click);
            // 
            // ImportCurrentTimeTable_btn
            // 
            this.ImportCurrentTimeTable_btn.Location = new System.Drawing.Point(554, 821);
            this.ImportCurrentTimeTable_btn.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.ImportCurrentTimeTable_btn.Name = "ImportCurrentTimeTable_btn";
            this.ImportCurrentTimeTable_btn.Size = new System.Drawing.Size(194, 81);
            this.ImportCurrentTimeTable_btn.TabIndex = 1;
            this.ImportCurrentTimeTable_btn.Text = "总表表头";
            this.ImportCurrentTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportCurrentTimeTable_btn.Click += new System.EventHandler(this.ImportCurrentTimeTable_btn_Click);
            // 
            // NewTimeTableFile_lbl
            // 
            this.NewTimeTableFile_lbl.AutoSize = true;
            this.NewTimeTableFile_lbl.Location = new System.Drawing.Point(214, 907);
            this.NewTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.NewTimeTableFile_lbl.Name = "NewTimeTableFile_lbl";
            this.NewTimeTableFile_lbl.Size = new System.Drawing.Size(0, 25);
            this.NewTimeTableFile_lbl.TabIndex = 2;
            // 
            // CurrentTimeTableFile_lbl
            // 
            this.CurrentTimeTableFile_lbl.AutoSize = true;
            this.CurrentTimeTableFile_lbl.Location = new System.Drawing.Point(654, 907);
            this.CurrentTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.CurrentTimeTableFile_lbl.Name = "CurrentTimeTableFile_lbl";
            this.CurrentTimeTableFile_lbl.Size = new System.Drawing.Size(0, 25);
            this.CurrentTimeTableFile_lbl.TabIndex = 3;
            // 
            // getTrains_btn
            // 
            this.getTrains_btn.Location = new System.Drawing.Point(1067, 821);
            this.getTrains_btn.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.getTrains_btn.Name = "getTrains_btn";
            this.getTrains_btn.Size = new System.Drawing.Size(194, 81);
            this.getTrains_btn.TabIndex = 4;
            this.getTrains_btn.Text = "处理";
            this.getTrains_btn.UseVisualStyleBackColor = true;
            this.getTrains_btn.Click += new System.EventHandler(this.getTrains_btn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(64, 169);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 25);
            this.label2.TabIndex = 8;
            this.label2.Text = "新时刻表车次";
            // 
            // newTrains_lv
            // 
            this.newTrains_lv.HideSelection = false;
            this.newTrains_lv.Location = new System.Drawing.Point(69, 212);
            this.newTrains_lv.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.newTrains_lv.Name = "newTrains_lv";
            this.newTrains_lv.Size = new System.Drawing.Size(1413, 569);
            this.newTrains_lv.TabIndex = 9;
            this.newTrains_lv.UseCompatibleStateImageBehavior = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(64, 19);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(192, 25);
            this.label3.TabIndex = 13;
            this.label3.Text = "时刻表表头提取的车站";
            // 
            // currentTimeTableStation_tb
            // 
            this.currentTimeTableStation_tb.Location = new System.Drawing.Point(69, 50);
            this.currentTimeTableStation_tb.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.currentTimeTableStation_tb.Name = "currentTimeTableStation_tb";
            this.currentTimeTableStation_tb.Size = new System.Drawing.Size(1413, 94);
            this.currentTimeTableStation_tb.TabIndex = 14;
            this.currentTimeTableStation_tb.Text = "";
            // 
            // trainCount_lb
            // 
            this.trainCount_lb.AutoSize = true;
            this.trainCount_lb.Location = new System.Drawing.Point(1354, 169);
            this.trainCount_lb.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.trainCount_lb.Name = "trainCount_lb";
            this.trainCount_lb.Size = new System.Drawing.Size(48, 25);
            this.trainCount_lb.TabIndex = 16;
            this.trainCount_lb.Text = "数量";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label1.Location = new System.Drawing.Point(560, 943);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(390, 25);
            this.label1.TabIndex = 17;
            this.label1.Text = "备注：从左至右点三个按钮选择文件，可以多选";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(259, 979);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(912, 25);
            this.label4.TabIndex = 18;
            this.label4.Text = "中间站接续列车需要自行添加，时刻表中列车运行顺序可能错误，徐兰场西向北列车未完全删除，注意查漏及对比";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label5.Location = new System.Drawing.Point(590, 1051);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(343, 25);
            this.label5.TabIndex = 19;
            this.label5.Text = "时刻表表头仅支持Excel 2003文件（*.xls）";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.OrangeRed;
            this.label6.Location = new System.Drawing.Point(1114, 1051);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(231, 33);
            this.label6.TabIndex = 20;
            this.label6.Text = "Build 20210518-yzcj";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label7.Location = new System.Drawing.Point(253, 1018);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(929, 25);
            this.label7.TabIndex = 21;
            this.label7.Text = "目前仅支持郑州东车站三场，可以在表头添加/删除中间站/线路所，注意表头格式（关键字：到达-股道-发出-通过）";
            // 
            // ImportDistributedTrainTimeTableFile_btn
            // 
            this.ImportDistributedTrainTimeTableFile_btn.Location = new System.Drawing.Point(756, 821);
            this.ImportDistributedTrainTimeTableFile_btn.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.ImportDistributedTrainTimeTableFile_btn.Name = "ImportDistributedTrainTimeTableFile_btn";
            this.ImportDistributedTrainTimeTableFile_btn.Size = new System.Drawing.Size(194, 81);
            this.ImportDistributedTrainTimeTableFile_btn.TabIndex = 22;
            this.ImportDistributedTrainTimeTableFile_btn.Text = "所有分表表头";
            this.ImportDistributedTrainTimeTableFile_btn.UseVisualStyleBackColor = true;
            this.ImportDistributedTrainTimeTableFile_btn.Click += new System.EventHandler(this.ImportDistributedTrainTimeTableFile_btn_Click);
            // 
            // DistributedTimeTableFile_lbl
            // 
            this.DistributedTimeTableFile_lbl.AutoSize = true;
            this.DistributedTimeTableFile_lbl.Location = new System.Drawing.Point(834, 907);
            this.DistributedTimeTableFile_lbl.Name = "DistributedTimeTableFile_lbl";
            this.DistributedTimeTableFile_lbl.Size = new System.Drawing.Size(18, 25);
            this.DistributedTimeTableFile_lbl.TabIndex = 23;
            this.DistributedTimeTableFile_lbl.Text = " ";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1518, 1126);
            this.Controls.Add(this.DistributedTimeTableFile_lbl);
            this.Controls.Add(this.ImportDistributedTrainTimeTableFile_btn);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.trainCount_lb);
            this.Controls.Add(this.currentTimeTableStation_tb);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.newTrains_lv);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.getTrains_btn);
            this.Controls.Add(this.CurrentTimeTableFile_lbl);
            this.Controls.Add(this.NewTimeTableFile_lbl);
            this.Controls.Add(this.ImportCurrentTimeTable_btn);
            this.Controls.Add(this.ImportNewTimeTable_btn);
            this.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.Name = "Main";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ImportNewTimeTable_btn;
        private System.Windows.Forms.Button ImportCurrentTimeTable_btn;
        private System.Windows.Forms.Label NewTimeTableFile_lbl;
        private System.Windows.Forms.Label CurrentTimeTableFile_lbl;
        private System.Windows.Forms.Button getTrains_btn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListView newTrains_lv;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox currentTimeTableStation_tb;
        private System.Windows.Forms.Label trainCount_lb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button ImportDistributedTrainTimeTableFile_btn;
        private System.Windows.Forms.Label DistributedTimeTableFile_lbl;
    }
}

