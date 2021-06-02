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
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.ImportDistributedTrainTimeTableFile_btn = new System.Windows.Forms.Button();
            this.DistributedTimeTableFile_lbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ImportNewTimeTable_btn
            // 
            this.ImportNewTimeTable_btn.Location = new System.Drawing.Point(263, 788);
            this.ImportNewTimeTable_btn.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.ImportNewTimeTable_btn.Name = "ImportNewTimeTable_btn";
            this.ImportNewTimeTable_btn.Size = new System.Drawing.Size(233, 78);
            this.ImportNewTimeTable_btn.TabIndex = 0;
            this.ImportNewTimeTable_btn.Text = "所有路局新时刻表";
            this.ImportNewTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportNewTimeTable_btn.Click += new System.EventHandler(this.ImportNewTimeTable_btn_Click);
            // 
            // ImportCurrentTimeTable_btn
            // 
            this.ImportCurrentTimeTable_btn.Location = new System.Drawing.Point(665, 788);
            this.ImportCurrentTimeTable_btn.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.ImportCurrentTimeTable_btn.Name = "ImportCurrentTimeTable_btn";
            this.ImportCurrentTimeTable_btn.Size = new System.Drawing.Size(233, 78);
            this.ImportCurrentTimeTable_btn.TabIndex = 1;
            this.ImportCurrentTimeTable_btn.Text = "总表表头";
            this.ImportCurrentTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportCurrentTimeTable_btn.Click += new System.EventHandler(this.ImportCurrentTimeTable_btn_Click);
            // 
            // NewTimeTableFile_lbl
            // 
            this.NewTimeTableFile_lbl.AutoSize = true;
            this.NewTimeTableFile_lbl.Location = new System.Drawing.Point(257, 871);
            this.NewTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.NewTimeTableFile_lbl.Name = "NewTimeTableFile_lbl";
            this.NewTimeTableFile_lbl.Size = new System.Drawing.Size(0, 24);
            this.NewTimeTableFile_lbl.TabIndex = 2;
            // 
            // CurrentTimeTableFile_lbl
            // 
            this.CurrentTimeTableFile_lbl.AutoSize = true;
            this.CurrentTimeTableFile_lbl.Location = new System.Drawing.Point(785, 871);
            this.CurrentTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.CurrentTimeTableFile_lbl.Name = "CurrentTimeTableFile_lbl";
            this.CurrentTimeTableFile_lbl.Size = new System.Drawing.Size(0, 24);
            this.CurrentTimeTableFile_lbl.TabIndex = 3;
            // 
            // getTrains_btn
            // 
            this.getTrains_btn.Location = new System.Drawing.Point(1280, 788);
            this.getTrains_btn.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getTrains_btn.Name = "getTrains_btn";
            this.getTrains_btn.Size = new System.Drawing.Size(233, 78);
            this.getTrains_btn.TabIndex = 4;
            this.getTrains_btn.Text = "处理";
            this.getTrains_btn.UseVisualStyleBackColor = true;
            this.getTrains_btn.Click += new System.EventHandler(this.getTrains_btn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(77, 162);
            this.label2.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(154, 24);
            this.label2.TabIndex = 8;
            this.label2.Text = "新时刻表车次";
            // 
            // newTrains_lv
            // 
            this.newTrains_lv.HideSelection = false;
            this.newTrains_lv.Location = new System.Drawing.Point(83, 204);
            this.newTrains_lv.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.newTrains_lv.Name = "newTrains_lv";
            this.newTrains_lv.Size = new System.Drawing.Size(1695, 546);
            this.newTrains_lv.TabIndex = 9;
            this.newTrains_lv.UseCompatibleStateImageBehavior = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(77, 18);
            this.label3.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(250, 24);
            this.label3.TabIndex = 13;
            this.label3.Text = "时刻表表头提取的车站";
            // 
            // currentTimeTableStation_tb
            // 
            this.currentTimeTableStation_tb.Location = new System.Drawing.Point(83, 48);
            this.currentTimeTableStation_tb.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.currentTimeTableStation_tb.Name = "currentTimeTableStation_tb";
            this.currentTimeTableStation_tb.Size = new System.Drawing.Size(1695, 90);
            this.currentTimeTableStation_tb.TabIndex = 14;
            this.currentTimeTableStation_tb.Text = "";
            // 
            // trainCount_lb
            // 
            this.trainCount_lb.AutoSize = true;
            this.trainCount_lb.Location = new System.Drawing.Point(1625, 162);
            this.trainCount_lb.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.trainCount_lb.Name = "trainCount_lb";
            this.trainCount_lb.Size = new System.Drawing.Size(58, 24);
            this.trainCount_lb.TabIndex = 16;
            this.trainCount_lb.Text = "数量";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label5.Location = new System.Drawing.Point(706, 965);
            this.label5.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(478, 24);
            this.label5.TabIndex = 19;
            this.label5.Text = "时刻表表头仅支持Excel 2003文件（*.xls）";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.OrangeRed;
            this.label6.Location = new System.Drawing.Point(1343, 941);
            this.label6.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(319, 33);
            this.label6.TabIndex = 20;
            this.label6.Text = "Build 20210518-yzcj";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label7.Location = new System.Drawing.Point(464, 941);
            this.label7.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(886, 24);
            this.label7.TabIndex = 21;
            this.label7.Text = "在表头添加/删除中间站/线路所，注意表头格式（关键字：到达-股道-发出-通过）";
            // 
            // ImportDistributedTrainTimeTableFile_btn
            // 
            this.ImportDistributedTrainTimeTableFile_btn.Location = new System.Drawing.Point(907, 788);
            this.ImportDistributedTrainTimeTableFile_btn.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.ImportDistributedTrainTimeTableFile_btn.Name = "ImportDistributedTrainTimeTableFile_btn";
            this.ImportDistributedTrainTimeTableFile_btn.Size = new System.Drawing.Size(233, 78);
            this.ImportDistributedTrainTimeTableFile_btn.TabIndex = 22;
            this.ImportDistributedTrainTimeTableFile_btn.Text = "所有分表表头";
            this.ImportDistributedTrainTimeTableFile_btn.UseVisualStyleBackColor = true;
            this.ImportDistributedTrainTimeTableFile_btn.Click += new System.EventHandler(this.ImportDistributedTrainTimeTableFile_btn_Click);
            // 
            // DistributedTimeTableFile_lbl
            // 
            this.DistributedTimeTableFile_lbl.AutoSize = true;
            this.DistributedTimeTableFile_lbl.Location = new System.Drawing.Point(1001, 871);
            this.DistributedTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.DistributedTimeTableFile_lbl.Name = "DistributedTimeTableFile_lbl";
            this.DistributedTimeTableFile_lbl.Size = new System.Drawing.Size(22, 24);
            this.DistributedTimeTableFile_lbl.TabIndex = 23;
            this.DistributedTimeTableFile_lbl.Text = " ";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1822, 1117);
            this.Controls.Add(this.DistributedTimeTableFile_lbl);
            this.Controls.Add(this.ImportDistributedTrainTimeTableFile_btn);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
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
            this.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
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
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button ImportDistributedTrainTimeTableFile_btn;
        private System.Windows.Forms.Label DistributedTimeTableFile_lbl;
    }
}

