namespace WordPrint
{
    partial class BatchPrint
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BatchPrint));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnPause = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.listBoxLog = new System.Windows.Forms.ListBox();
            this.btnBreak = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.pgrShow = new System.Windows.Forms.ProgressBar();
            this.btnDownLoad = new System.Windows.Forms.Button();
            this.lblFirst = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.listBoxFile = new System.Windows.Forms.ListBox();
            this.btnLoadFile = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cmbPrinter = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtPrintName = new System.Windows.Forms.TextBox();
            this.txtRestTime = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtCount = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblMsg = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox5);
            this.groupBox1.Controls.Add(this.groupBox4);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Location = new System.Drawing.Point(12, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(749, 523);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "打印控制";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnClear);
            this.groupBox5.Controls.Add(this.btnPause);
            this.groupBox5.Controls.Add(this.label7);
            this.groupBox5.Controls.Add(this.listBoxLog);
            this.groupBox5.Controls.Add(this.btnBreak);
            this.groupBox5.Controls.Add(this.btnStart);
            this.groupBox5.Controls.Add(this.label3);
            this.groupBox5.Location = new System.Drawing.Point(323, 357);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(419, 158);
            this.groupBox5.TabIndex = 4;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "第四步";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(329, 123);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(63, 23);
            this.btnClear.TabIndex = 6;
            this.btnClear.Text = "清空状态";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnPause
            // 
            this.btnPause.Location = new System.Drawing.Point(284, 49);
            this.btnPause.Name = "btnPause";
            this.btnPause.Size = new System.Drawing.Size(62, 23);
            this.btnPause.TabIndex = 5;
            this.btnPause.Text = "暂停";
            this.btnPause.UseVisualStyleBackColor = true;
            this.btnPause.Visible = false;
            this.btnPause.Click += new System.EventHandler(this.btnPause_Click);
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(20, 24);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(25, 83);
            this.label7.TabIndex = 4;
            this.label7.Text = "运行状态";
            // 
            // listBoxLog
            // 
            this.listBoxLog.FormattingEnabled = true;
            this.listBoxLog.ItemHeight = 12;
            this.listBoxLog.Location = new System.Drawing.Point(51, 19);
            this.listBoxLog.Name = "listBoxLog";
            this.listBoxLog.Size = new System.Drawing.Size(362, 88);
            this.listBoxLog.TabIndex = 3;
            // 
            // btnBreak
            // 
            this.btnBreak.Enabled = false;
            this.btnBreak.Location = new System.Drawing.Point(263, 123);
            this.btnBreak.Name = "btnBreak";
            this.btnBreak.Size = new System.Drawing.Size(60, 23);
            this.btnBreak.TabIndex = 2;
            this.btnBreak.Text = "中断";
            this.btnBreak.UseVisualStyleBackColor = true;
            this.btnBreak.Visible = false;
            this.btnBreak.Click += new System.EventHandler(this.btnBreak_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(193, 123);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(60, 23);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "启动";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 114);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(173, 36);
            this.label3.TabIndex = 1;
            this.label3.Text = "点击\"开始打印\"按钮，调用打印\r\n机进行打印。（打印之前请先检\r\n查打印机设备）";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.pgrShow);
            this.groupBox4.Controls.Add(this.btnDownLoad);
            this.groupBox4.Controls.Add(this.lblFirst);
            this.groupBox4.Location = new System.Drawing.Point(6, 20);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(737, 57);
            this.groupBox4.TabIndex = 7;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "第一步";
            // 
            // pgrShow
            // 
            this.pgrShow.Location = new System.Drawing.Point(317, 25);
            this.pgrShow.Name = "pgrShow";
            this.pgrShow.Size = new System.Drawing.Size(311, 14);
            this.pgrShow.TabIndex = 1;
            this.pgrShow.Visible = false;
            // 
            // btnDownLoad
            // 
            this.btnDownLoad.Location = new System.Drawing.Point(634, 20);
            this.btnDownLoad.Name = "btnDownLoad";
            this.btnDownLoad.Size = new System.Drawing.Size(75, 23);
            this.btnDownLoad.TabIndex = 0;
            this.btnDownLoad.Text = "下载";
            this.btnDownLoad.UseVisualStyleBackColor = true;
            this.btnDownLoad.Click += new System.EventHandler(this.btnDownLoad_Click);
            // 
            // lblFirst
            // 
            this.lblFirst.AutoSize = true;
            this.lblFirst.Location = new System.Drawing.Point(13, 25);
            this.lblFirst.Name = "lblFirst";
            this.lblFirst.Size = new System.Drawing.Size(317, 12);
            this.lblFirst.TabIndex = 1;
            this.lblFirst.Text = "点击\"进入文件夹\"按钮，把要打印的文件放到该文件夹下。";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.listBoxFile);
            this.groupBox3.Controls.Add(this.btnLoadFile);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Location = new System.Drawing.Point(6, 83);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(736, 268);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "第二步";
            // 
            // listBoxFile
            // 
            this.listBoxFile.FormattingEnabled = true;
            this.listBoxFile.ItemHeight = 12;
            this.listBoxFile.Location = new System.Drawing.Point(14, 54);
            this.listBoxFile.Name = "listBoxFile";
            this.listBoxFile.Size = new System.Drawing.Size(716, 208);
            this.listBoxFile.TabIndex = 3;
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.Location = new System.Drawing.Point(633, 20);
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.Size = new System.Drawing.Size(75, 23);
            this.btnLoadFile.TabIndex = 0;
            this.btnLoadFile.Text = "加载文件";
            this.btnLoadFile.UseVisualStyleBackColor = true;
            this.btnLoadFile.Click += new System.EventHandler(this.btnLoadFile_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(401, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "点击\"加载文件\"按钮，把文件夹中的文件加载到列表中。检查文件正确性。";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cmbPrinter);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.txtPrintName);
            this.groupBox2.Controls.Add(this.txtRestTime);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.txtCount);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(7, 357);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(310, 160);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "第三步";
            // 
            // cmbPrinter
            // 
            this.cmbPrinter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPrinter.FormattingEnabled = true;
            this.cmbPrinter.Location = new System.Drawing.Point(22, 132);
            this.cmbPrinter.Name = "cmbPrinter";
            this.cmbPrinter.Size = new System.Drawing.Size(276, 20);
            this.cmbPrinter.TabIndex = 9;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(233, 78);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(17, 12);
            this.label10.TabIndex = 8;
            this.label10.Text = "秒";
            // 
            // txtPrintName
            // 
            this.txtPrintName.Location = new System.Drawing.Point(127, 106);
            this.txtPrintName.Multiline = true;
            this.txtPrintName.Name = "txtPrintName";
            this.txtPrintName.Size = new System.Drawing.Size(171, 20);
            this.txtPrintName.TabIndex = 7;
            this.txtPrintName.Text = "Foxit Reader PDF Printer";
            this.txtPrintName.Visible = false;
            // 
            // txtRestTime
            // 
            this.txtRestTime.Location = new System.Drawing.Point(127, 75);
            this.txtRestTime.Name = "txtRestTime";
            this.txtRestTime.Size = new System.Drawing.Size(100, 21);
            this.txtRestTime.TabIndex = 7;
            this.txtRestTime.Text = "20";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(20, 109);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(77, 12);
            this.label8.TabIndex = 6;
            this.label8.Text = "选择打印机：";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(20, 78);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(101, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "每次间歇的时间：";
            // 
            // txtCount
            // 
            this.txtCount.Location = new System.Drawing.Point(127, 43);
            this.txtCount.Name = "txtCount";
            this.txtCount.Size = new System.Drawing.Size(100, 21);
            this.txtCount.TabIndex = 5;
            this.txtCount.Text = "10";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(233, 49);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(65, 12);
            this.label9.TabIndex = 4;
            this.label9.Text = "份休息一次";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(69, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 4;
            this.label5.Text = "每打印";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 12);
            this.label4.TabIndex = 1;
            this.label4.Text = "设置打印的参数";
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(19, 532);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(0, 12);
            this.lblMsg.TabIndex = 1;
            // 
            // BatchPrint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(773, 558);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "BatchPrint";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "批量打印工具";
            this.groupBox1.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblFirst;
        private System.Windows.Forms.Button btnDownLoad;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnLoadFile;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ListBox listBoxFile;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtRestTime;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtCount;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button btnBreak;
        private System.Windows.Forms.ListBox listBoxLog;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtPrintName;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnPause;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cmbPrinter;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ProgressBar pgrShow;
        private System.Windows.Forms.Label lblMsg;
    }
}

