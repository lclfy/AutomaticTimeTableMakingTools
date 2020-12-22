﻿namespace AutomaticTimeTableMakingTools
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
            this.SuspendLayout();
            // 
            // ImportNewTimeTable_btn
            // 
            this.ImportNewTimeTable_btn.Location = new System.Drawing.Point(240, 690);
            this.ImportNewTimeTable_btn.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.ImportNewTimeTable_btn.Name = "ImportNewTimeTable_btn";
            this.ImportNewTimeTable_btn.Size = new System.Drawing.Size(215, 68);
            this.ImportNewTimeTable_btn.TabIndex = 0;
            this.ImportNewTimeTable_btn.Text = "所有路局新时刻表";
            this.ImportNewTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportNewTimeTable_btn.Click += new System.EventHandler(this.ImportNewTimeTable_btn_Click);
            // 
            // ImportCurrentTimeTable_btn
            // 
            this.ImportCurrentTimeTable_btn.Location = new System.Drawing.Point(724, 690);
            this.ImportCurrentTimeTable_btn.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.ImportCurrentTimeTable_btn.Name = "ImportCurrentTimeTable_btn";
            this.ImportCurrentTimeTable_btn.Size = new System.Drawing.Size(215, 68);
            this.ImportCurrentTimeTable_btn.TabIndex = 1;
            this.ImportCurrentTimeTable_btn.Text = "要修改的表头";
            this.ImportCurrentTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportCurrentTimeTable_btn.Click += new System.EventHandler(this.ImportCurrentTimeTable_btn_Click);
            // 
            // NewTimeTableFile_lbl
            // 
            this.NewTimeTableFile_lbl.AutoSize = true;
            this.NewTimeTableFile_lbl.Location = new System.Drawing.Point(237, 763);
            this.NewTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.NewTimeTableFile_lbl.Name = "NewTimeTableFile_lbl";
            this.NewTimeTableFile_lbl.Size = new System.Drawing.Size(0, 21);
            this.NewTimeTableFile_lbl.TabIndex = 2;
            // 
            // CurrentTimeTableFile_lbl
            // 
            this.CurrentTimeTableFile_lbl.AutoSize = true;
            this.CurrentTimeTableFile_lbl.Location = new System.Drawing.Point(721, 763);
            this.CurrentTimeTableFile_lbl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.CurrentTimeTableFile_lbl.Name = "CurrentTimeTableFile_lbl";
            this.CurrentTimeTableFile_lbl.Size = new System.Drawing.Size(0, 21);
            this.CurrentTimeTableFile_lbl.TabIndex = 3;
            // 
            // getTrains_btn
            // 
            this.getTrains_btn.Location = new System.Drawing.Point(1173, 690);
            this.getTrains_btn.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.getTrains_btn.Name = "getTrains_btn";
            this.getTrains_btn.Size = new System.Drawing.Size(215, 68);
            this.getTrains_btn.TabIndex = 4;
            this.getTrains_btn.Text = "处理";
            this.getTrains_btn.UseVisualStyleBackColor = true;
            this.getTrains_btn.Click += new System.EventHandler(this.getTrains_btn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(72, 142);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 21);
            this.label2.TabIndex = 8;
            this.label2.Text = "新时刻表车次";
            // 
            // newTrains_lv
            // 
            this.newTrains_lv.HideSelection = false;
            this.newTrains_lv.Location = new System.Drawing.Point(75, 178);
            this.newTrains_lv.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.newTrains_lv.Name = "newTrains_lv";
            this.newTrains_lv.Size = new System.Drawing.Size(1553, 478);
            this.newTrains_lv.TabIndex = 9;
            this.newTrains_lv.UseCompatibleStateImageBehavior = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(72, 16);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(220, 21);
            this.label3.TabIndex = 13;
            this.label3.Text = "时刻表表头提取的车站";
            // 
            // currentTimeTableStation_tb
            // 
            this.currentTimeTableStation_tb.Location = new System.Drawing.Point(75, 42);
            this.currentTimeTableStation_tb.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.currentTimeTableStation_tb.Name = "currentTimeTableStation_tb";
            this.currentTimeTableStation_tb.Size = new System.Drawing.Size(1553, 79);
            this.currentTimeTableStation_tb.TabIndex = 14;
            this.currentTimeTableStation_tb.Text = "";
            // 
            // trainCount_lb
            // 
            this.trainCount_lb.AutoSize = true;
            this.trainCount_lb.Location = new System.Drawing.Point(1491, 142);
            this.trainCount_lb.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.trainCount_lb.Name = "trainCount_lb";
            this.trainCount_lb.Size = new System.Drawing.Size(52, 21);
            this.trainCount_lb.TabIndex = 16;
            this.trainCount_lb.Text = "数量";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label1.Location = new System.Drawing.Point(616, 791);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(451, 21);
            this.label1.TabIndex = 17;
            this.label1.Text = "备注：从左至右点三个按钮选择文件，可以多选";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(284, 822);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(1060, 21);
            this.label4.TabIndex = 18;
            this.label4.Text = "中间站接续列车需要自行添加，时刻表中列车运行顺序可能错误，徐兰场西向北列车未完全删除，注意查漏及对比";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label5.Location = new System.Drawing.Point(649, 884);
            this.label5.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(427, 21);
            this.label5.TabIndex = 19;
            this.label5.Text = "时刻表表头仅支持Excel 2003文件（*.xls）";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
<<<<<<< HEAD:AutomaticTimeTableMakingTools/Main.Designer.cs
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.OrangeRed;
            this.label6.Location = new System.Drawing.Point(1227, 884);
            this.label6.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(208, 28);
            this.label6.TabIndex = 20;
            this.label6.Text = "Build 20191208";
=======
            this.label6.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label6.Location = new System.Drawing.Point(669, 505);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 12);
            this.label6.TabIndex = 20;
            this.label6.Text = "Build 20180320";
>>>>>>> parent of 707657c... 181205:Main.Designer.cs
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label7.Location = new System.Drawing.Point(279, 854);
            this.label7.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(1094, 21);
            this.label7.TabIndex = 21;
            this.label7.Text = "目前仅支持郑州东车站三场，可以在表头添加/删除中间站/线路所，注意表头格式（关键字：到达-股道-发出-通过）";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
<<<<<<< HEAD:AutomaticTimeTableMakingTools/Main.Designer.cs
            this.ClientSize = new System.Drawing.Size(1714, 934);
=======
            this.ClientSize = new System.Drawing.Size(935, 526);
>>>>>>> parent of 707657c... 181205:Main.Designer.cs
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
            this.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
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
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox currentTimeTableStation_tb;
        private System.Windows.Forms.Label trainCount_lb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
    }
}
