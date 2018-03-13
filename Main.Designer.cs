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
            this.SuspendLayout();
            // 
            // ImportNewTimeTable_btn
            // 
            this.ImportNewTimeTable_btn.Location = new System.Drawing.Point(131, 394);
            this.ImportNewTimeTable_btn.Name = "ImportNewTimeTable_btn";
            this.ImportNewTimeTable_btn.Size = new System.Drawing.Size(117, 39);
            this.ImportNewTimeTable_btn.TabIndex = 0;
            this.ImportNewTimeTable_btn.Text = "导入新时刻表";
            this.ImportNewTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportNewTimeTable_btn.Click += new System.EventHandler(this.ImportNewTimeTable_btn_Click);
            // 
            // ImportCurrentTimeTable_btn
            // 
            this.ImportCurrentTimeTable_btn.Location = new System.Drawing.Point(395, 394);
            this.ImportCurrentTimeTable_btn.Name = "ImportCurrentTimeTable_btn";
            this.ImportCurrentTimeTable_btn.Size = new System.Drawing.Size(117, 39);
            this.ImportCurrentTimeTable_btn.TabIndex = 1;
            this.ImportCurrentTimeTable_btn.Text = "要修改的时刻表";
            this.ImportCurrentTimeTable_btn.UseVisualStyleBackColor = true;
            this.ImportCurrentTimeTable_btn.Click += new System.EventHandler(this.ImportCurrentTimeTable_btn_Click);
            // 
            // NewTimeTableFile_lbl
            // 
            this.NewTimeTableFile_lbl.AutoSize = true;
            this.NewTimeTableFile_lbl.Location = new System.Drawing.Point(129, 436);
            this.NewTimeTableFile_lbl.Name = "NewTimeTableFile_lbl";
            this.NewTimeTableFile_lbl.Size = new System.Drawing.Size(0, 12);
            this.NewTimeTableFile_lbl.TabIndex = 2;
            // 
            // CurrentTimeTableFile_lbl
            // 
            this.CurrentTimeTableFile_lbl.AutoSize = true;
            this.CurrentTimeTableFile_lbl.Location = new System.Drawing.Point(393, 436);
            this.CurrentTimeTableFile_lbl.Name = "CurrentTimeTableFile_lbl";
            this.CurrentTimeTableFile_lbl.Size = new System.Drawing.Size(0, 12);
            this.CurrentTimeTableFile_lbl.TabIndex = 3;
            // 
            // getTrains_btn
            // 
            this.getTrains_btn.Location = new System.Drawing.Point(640, 394);
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
            this.label2.Location = new System.Drawing.Point(39, 81);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "新时刻表车次";
            // 
            // newTrains_lv
            // 
            this.newTrains_lv.Location = new System.Drawing.Point(41, 102);
            this.newTrains_lv.Name = "newTrains_lv";
            this.newTrains_lv.Size = new System.Drawing.Size(849, 275);
            this.newTrains_lv.TabIndex = 9;
            this.newTrains_lv.UseCompatibleStateImageBehavior = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(39, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 12);
            this.label3.TabIndex = 13;
            this.label3.Text = "时刻表表头提取的车站";
            // 
            // currentTimeTableStation_tb
            // 
            this.currentTimeTableStation_tb.Location = new System.Drawing.Point(41, 24);
            this.currentTimeTableStation_tb.Name = "currentTimeTableStation_tb";
            this.currentTimeTableStation_tb.Size = new System.Drawing.Size(849, 47);
            this.currentTimeTableStation_tb.TabIndex = 14;
            this.currentTimeTableStation_tb.Text = "";
            // 
            // trainCount_lb
            // 
            this.trainCount_lb.AutoSize = true;
            this.trainCount_lb.Location = new System.Drawing.Point(813, 81);
            this.trainCount_lb.Name = "trainCount_lb";
            this.trainCount_lb.Size = new System.Drawing.Size(29, 12);
            this.trainCount_lb.TabIndex = 16;
            this.trainCount_lb.Text = "数量";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(935, 469);
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
    }
}

