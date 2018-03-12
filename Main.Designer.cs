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
            this.mainStationName_tb = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.currentTimeTableStation_tb = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // ImportNewTimeTable_btn
            // 
            this.ImportNewTimeTable_btn.Location = new System.Drawing.Point(50, 21);
            this.ImportNewTimeTable_btn.Name = "ImportNewTimeTable_btn";
            this.ImportNewTimeTable_btn.Size = new System.Drawing.Size(117, 39);
            this.ImportNewTimeTable_btn.TabIndex = 0;
            this.ImportNewTimeTable_btn.Text = "导入新时刻表";
            this.ImportNewTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportNewTimeTable_btn.Click += new System.EventHandler(this.ImportNewTimeTable_btn_Click);
            // 
            // ImportCurrentTimeTable_btn
            // 
            this.ImportCurrentTimeTable_btn.Location = new System.Drawing.Point(50, 196);
            this.ImportCurrentTimeTable_btn.Name = "ImportCurrentTimeTable_btn";
            this.ImportCurrentTimeTable_btn.Size = new System.Drawing.Size(117, 39);
            this.ImportCurrentTimeTable_btn.TabIndex = 1;
            this.ImportCurrentTimeTable_btn.Text = "导入本站时刻表";
            this.ImportCurrentTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportCurrentTimeTable_btn.Click += new System.EventHandler(this.ImportCurrentTimeTable_btn_Click);
            // 
            // NewTimeTableFile_lbl
            // 
            this.NewTimeTableFile_lbl.AutoSize = true;
            this.NewTimeTableFile_lbl.Location = new System.Drawing.Point(48, 83);
            this.NewTimeTableFile_lbl.Name = "NewTimeTableFile_lbl";
            this.NewTimeTableFile_lbl.Size = new System.Drawing.Size(125, 12);
            this.NewTimeTableFile_lbl.TabIndex = 2;
            this.NewTimeTableFile_lbl.Text = "newTimeTableFile_lbl";
            // 
            // CurrentTimeTableFile_lbl
            // 
            this.CurrentTimeTableFile_lbl.AutoSize = true;
            this.CurrentTimeTableFile_lbl.Location = new System.Drawing.Point(48, 253);
            this.CurrentTimeTableFile_lbl.Name = "CurrentTimeTableFile_lbl";
            this.CurrentTimeTableFile_lbl.Size = new System.Drawing.Size(41, 12);
            this.CurrentTimeTableFile_lbl.TabIndex = 3;
            this.CurrentTimeTableFile_lbl.Text = "label2";
            // 
            // getTrains_btn
            // 
            this.getTrains_btn.Location = new System.Drawing.Point(50, 349);
            this.getTrains_btn.Name = "getTrains_btn";
            this.getTrains_btn.Size = new System.Drawing.Size(117, 39);
            this.getTrains_btn.TabIndex = 4;
            this.getTrains_btn.Text = "读取车次信息";
            this.getTrains_btn.UseVisualStyleBackColor = true;
            this.getTrains_btn.Click += new System.EventHandler(this.getTrains_btn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(207, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "新时刻表车次";
            // 
            // newTrains_lv
            // 
            this.newTrains_lv.Location = new System.Drawing.Point(209, 97);
            this.newTrains_lv.Name = "newTrains_lv";
            this.newTrains_lv.Size = new System.Drawing.Size(849, 256);
            this.newTrains_lv.TabIndex = 9;
            this.newTrains_lv.UseCompatibleStateImageBehavior = false;
            // 
            // mainStationName_tb
            // 
            this.mainStationName_tb.Location = new System.Drawing.Point(209, 388);
            this.mainStationName_tb.Name = "mainStationName_tb";
            this.mainStationName_tb.Size = new System.Drawing.Size(126, 21);
            this.mainStationName_tb.TabIndex = 10;
            this.mainStationName_tb.Text = "郑州东京广场/徐兰场";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(207, 369);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 11;
            this.label1.Text = "主站站名";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(207, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 13;
            this.label3.Text = "当前时刻表车站";
            // 
            // currentTimeTableStation_tb
            // 
            this.currentTimeTableStation_tb.Location = new System.Drawing.Point(209, 32);
            this.currentTimeTableStation_tb.Name = "currentTimeTableStation_tb";
            this.currentTimeTableStation_tb.Size = new System.Drawing.Size(849, 47);
            this.currentTimeTableStation_tb.TabIndex = 14;
            this.currentTimeTableStation_tb.Text = "";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1116, 459);
            this.Controls.Add(this.currentTimeTableStation_tb);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.mainStationName_tb);
            this.Controls.Add(this.newTrains_lv);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.getTrains_btn);
            this.Controls.Add(this.CurrentTimeTableFile_lbl);
            this.Controls.Add(this.NewTimeTableFile_lbl);
            this.Controls.Add(this.ImportCurrentTimeTable_btn);
            this.Controls.Add(this.ImportNewTimeTable_btn);
            this.Name = "Main";
            this.Text = "Form1";
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
        private System.Windows.Forms.TextBox mainStationName_tb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox currentTimeTableStation_tb;
    }
}

