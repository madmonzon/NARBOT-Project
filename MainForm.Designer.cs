using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace NARBOT
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.MidmonthGB = new System.Windows.Forms.GroupBox();
            this.MidMonthLB = new System.Windows.Forms.CheckedListBox();
            this.AgingUtilchk = new System.Windows.Forms.CheckBox();
            this.SplitRunCHK = new System.Windows.Forms.CheckBox();
            this.Closebtn = new System.Windows.Forms.Button();
            this.BOT_ICON = new System.Windows.Forms.NotifyIcon(this.components);
            this.ROToolchk = new System.Windows.Forms.CheckBox();
            this.SundaysISNCB = new System.Windows.Forms.CheckBox();
            this.BIWeeklyISNCB = new System.Windows.Forms.CheckBox();
            this.SaturdaysGB = new System.Windows.Forms.GroupBox();
            this.SaturdaysLB = new System.Windows.Forms.CheckedListBox();
            this.EOMGB = new System.Windows.Forms.GroupBox();
            this.EOMLB = new System.Windows.Forms.CheckedListBox();
            this.EOMISNCB = new System.Windows.Forms.CheckBox();
            this.MidMonthISNCB = new System.Windows.Forms.CheckBox();
            this.RMannuallyCB = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.SundaysLB = new System.Windows.Forms.CheckedListBox();
            this.BIWeeklyGb = new System.Windows.Forms.GroupBox();
            this.BiWeeklyLB = new System.Windows.Forms.CheckedListBox();
            this.Messageslb = new System.Windows.Forms.ListBox();
            this.BOT_timer = new System.Windows.Forms.Timer(this.components);
            this.BTNSGB = new System.Windows.Forms.GroupBox();
            this.UnCheckbtn = new System.Windows.Forms.Button();
            this.CheckAll = new System.Windows.Forms.Button();
            this.TimeShowlbl = new System.Windows.Forms.Label();
            this.TimeTXTlbl = new System.Windows.Forms.Label();
            this.DateShowlbl = new System.Windows.Forms.Label();
            this.DateTXTlbl = new System.Windows.Forms.Label();
            this.Run_button = new System.Windows.Forms.Button();
            this.WeekListBox = new System.Windows.Forms.CheckedListBox();
            this.RunbtnGB = new System.Windows.Forms.GroupBox();
            this.ProcessesGB = new System.Windows.Forms.GroupBox();
            this.chkBalUtil = new System.Windows.Forms.CheckBox();
            this.MidmonthGB.SuspendLayout();
            this.SaturdaysGB.SuspendLayout();
            this.EOMGB.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.BIWeeklyGb.SuspendLayout();
            this.BTNSGB.SuspendLayout();
            this.RunbtnGB.SuspendLayout();
            this.ProcessesGB.SuspendLayout();
            this.SuspendLayout();
            // 
            // MidmonthGB
            // 
            this.MidmonthGB.Controls.Add(this.MidMonthLB);
            this.MidmonthGB.Location = new System.Drawing.Point(12, 421);
            this.MidmonthGB.Name = "MidmonthGB";
            this.MidmonthGB.Size = new System.Drawing.Size(282, 91);
            this.MidmonthGB.TabIndex = 42;
            this.MidmonthGB.TabStop = false;
            this.MidmonthGB.Text = "15th of Month";
            // 
            // MidMonthLB
            // 
            this.MidMonthLB.CheckOnClick = true;
            this.MidMonthLB.ColumnWidth = 284;
            this.MidMonthLB.FormattingEnabled = true;
            this.MidMonthLB.Items.AddRange(new object[] {
            "ISN Printer Mid-Month - Other (1)",
            "ISN Printer Mid-Month - Ledger (2)",
            "ISN Printer Mid-Month - NonLedger (3)",
            "ISN Printer Mid-Month - Agencies (4)"});
            this.MidMonthLB.Location = new System.Drawing.Point(8, 19);
            this.MidMonthLB.MultiColumn = true;
            this.MidMonthLB.Name = "MidMonthLB";
            this.MidMonthLB.Size = new System.Drawing.Size(268, 64);
            this.MidMonthLB.TabIndex = 15;
            // 
            // AgingUtilchk
            // 
            this.AgingUtilchk.AutoSize = true;
            this.AgingUtilchk.Location = new System.Drawing.Point(191, 518);
            this.AgingUtilchk.Name = "AgingUtilchk";
            this.AgingUtilchk.Size = new System.Drawing.Size(112, 17);
            this.AgingUtilchk.TabIndex = 50;
            this.AgingUtilchk.Text = "Stop Aging Utility?";
            this.AgingUtilchk.UseVisualStyleBackColor = true;
            // 
            // SplitRunCHK
            // 
            this.SplitRunCHK.AutoSize = true;
            this.SplitRunCHK.Enabled = false;
            this.SplitRunCHK.Location = new System.Drawing.Point(17, 541);
            this.SplitRunCHK.Name = "SplitRunCHK";
            this.SplitRunCHK.Size = new System.Drawing.Size(69, 17);
            this.SplitRunCHK.TabIndex = 47;
            this.SplitRunCHK.Text = "Split Run";
            this.SplitRunCHK.UseVisualStyleBackColor = true;
            // 
            // Closebtn
            // 
            this.Closebtn.Location = new System.Drawing.Point(555, 11);
            this.Closebtn.Name = "Closebtn";
            this.Closebtn.Size = new System.Drawing.Size(42, 23);
            this.Closebtn.TabIndex = 33;
            this.Closebtn.Text = "Close";
            this.Closebtn.UseVisualStyleBackColor = true;
            this.Closebtn.Click += new System.EventHandler(this.Closebtn_Click);
            // 
            // BOT_ICON
            // 
            this.BOT_ICON.Icon = ((System.Drawing.Icon)(resources.GetObject("BOT_ICON.Icon")));
            this.BOT_ICON.Text = "NARBOT";
            this.BOT_ICON.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.BOT_MouseDoubleClick);
            // 
            // ROToolchk
            // 
            this.ROToolchk.AutoSize = true;
            this.ROToolchk.Location = new System.Drawing.Point(191, 541);
            this.ROToolchk.Name = "ROToolchk";
            this.ROToolchk.Size = new System.Drawing.Size(121, 17);
            this.ROToolchk.TabIndex = 49;
            this.ROToolchk.Text = "Stop ReOpen Tool?";
            this.ROToolchk.UseVisualStyleBackColor = true;
            // 
            // SundaysISNCB
            // 
            this.SundaysISNCB.AutoSize = true;
            this.SundaysISNCB.Location = new System.Drawing.Point(463, 518);
            this.SundaysISNCB.Name = "SundaysISNCB";
            this.SundaysISNCB.Size = new System.Drawing.Size(119, 17);
            this.SundaysISNCB.TabIndex = 48;
            this.SundaysISNCB.Text = "Stop Sundays ISN?";
            this.SundaysISNCB.UseVisualStyleBackColor = true;
            // 
            // BIWeeklyISNCB
            // 
            this.BIWeeklyISNCB.AutoSize = true;
            this.BIWeeklyISNCB.Location = new System.Drawing.Point(324, 541);
            this.BIWeeklyISNCB.Name = "BIWeeklyISNCB";
            this.BIWeeklyISNCB.Size = new System.Drawing.Size(127, 17);
            this.BIWeeklyISNCB.TabIndex = 45;
            this.BIWeeklyISNCB.Text = "Stop BI-Weekly ISN?";
            this.BIWeeklyISNCB.UseVisualStyleBackColor = true;
            // 
            // SaturdaysGB
            // 
            this.SaturdaysGB.Controls.Add(this.SaturdaysLB);
            this.SaturdaysGB.Location = new System.Drawing.Point(12, 362);
            this.SaturdaysGB.Name = "SaturdaysGB";
            this.SaturdaysGB.Size = new System.Drawing.Size(282, 60);
            this.SaturdaysGB.TabIndex = 41;
            this.SaturdaysGB.TabStop = false;
            this.SaturdaysGB.Text = "Saturdays";
            // 
            // SaturdaysLB
            // 
            this.SaturdaysLB.CheckOnClick = true;
            this.SaturdaysLB.ColumnWidth = 284;
            this.SaturdaysLB.FormattingEnabled = true;
            this.SaturdaysLB.Items.AddRange(new object[] {
            "Balancing Utility (1)",
            "Oracle Re-open Process (2)"});
            this.SaturdaysLB.Location = new System.Drawing.Point(5, 19);
            this.SaturdaysLB.MultiColumn = true;
            this.SaturdaysLB.Name = "SaturdaysLB";
            this.SaturdaysLB.Size = new System.Drawing.Size(271, 34);
            this.SaturdaysLB.TabIndex = 6;
            // 
            // EOMGB
            // 
            this.EOMGB.Controls.Add(this.EOMLB);
            this.EOMGB.Location = new System.Drawing.Point(300, 408);
            this.EOMGB.Name = "EOMGB";
            this.EOMGB.Size = new System.Drawing.Size(297, 104);
            this.EOMGB.TabIndex = 40;
            this.EOMGB.TabStop = false;
            this.EOMGB.Text = "End of Month";
            // 
            // EOMLB
            // 
            this.EOMLB.CheckOnClick = true;
            this.EOMLB.ColumnWidth = 284;
            this.EOMLB.FormattingEnabled = true;
            this.EOMLB.Items.AddRange(new object[] {
            "Finance Manager (1)",
            "ISN Printer EOM- Other (2)",
            "ISN Printer EOM- Agencies (3)",
            "ISN Printer EOM- Ledger (4)",
            "ISN Printer EOM- NonLedger (5)"});
            this.EOMLB.Location = new System.Drawing.Point(8, 19);
            this.EOMLB.MultiColumn = true;
            this.EOMLB.Name = "EOMLB";
            this.EOMLB.Size = new System.Drawing.Size(283, 79);
            this.EOMLB.TabIndex = 15;
            // 
            // EOMISNCB
            // 
            this.EOMISNCB.AutoSize = true;
            this.EOMISNCB.Location = new System.Drawing.Point(324, 518);
            this.EOMISNCB.Name = "EOMISNCB";
            this.EOMISNCB.Size = new System.Drawing.Size(102, 17);
            this.EOMISNCB.TabIndex = 46;
            this.EOMISNCB.Text = "Stop EOM ISN?";
            this.EOMISNCB.UseVisualStyleBackColor = true;
            // 
            // MidMonthISNCB
            // 
            this.MidMonthISNCB.AutoSize = true;
            this.MidMonthISNCB.Location = new System.Drawing.Point(463, 541);
            this.MidMonthISNCB.Name = "MidMonthISNCB";
            this.MidMonthISNCB.Size = new System.Drawing.Size(128, 17);
            this.MidMonthISNCB.TabIndex = 44;
            this.MidMonthISNCB.Text = "Stop Mid-Month ISN?";
            this.MidMonthISNCB.UseVisualStyleBackColor = true;
            // 
            // RMannuallyCB
            // 
            this.RMannuallyCB.AutoSize = true;
            this.RMannuallyCB.Location = new System.Drawing.Point(17, 518);
            this.RMannuallyCB.Name = "RMannuallyCB";
            this.RMannuallyCB.Size = new System.Drawing.Size(168, 17);
            this.RMannuallyCB.TabIndex = 43;
            this.RMannuallyCB.Text = "Run Manually Un-Scheduled?";
            this.RMannuallyCB.UseVisualStyleBackColor = true;
            this.RMannuallyCB.CheckedChanged += new System.EventHandler(this.RMannuallyCB_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.SundaysLB);
            this.groupBox1.Location = new System.Drawing.Point(300, 299);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(297, 103);
            this.groupBox1.TabIndex = 39;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Sundays";
            // 
            // SundaysLB
            // 
            this.SundaysLB.CheckOnClick = true;
            this.SundaysLB.ColumnWidth = 284;
            this.SundaysLB.FormattingEnabled = true;
            this.SundaysLB.Items.AddRange(new object[] {
            "ISN Printer Sundays - Agencies (1)",
            "Run Statement  Report - Agencies (2)",
            "ISN Printer Sundays - Ledger (3)",
            "Run Statement  Report - Ledger (4)",
            "Aging Utility - Weekly (5)"});
            this.SundaysLB.Location = new System.Drawing.Point(5, 19);
            this.SundaysLB.MultiColumn = true;
            this.SundaysLB.Name = "SundaysLB";
            this.SundaysLB.Size = new System.Drawing.Size(286, 79);
            this.SundaysLB.TabIndex = 6;
            // 
            // BIWeeklyGb
            // 
            this.BIWeeklyGb.Controls.Add(this.BiWeeklyLB);
            this.BIWeeklyGb.Location = new System.Drawing.Point(12, 299);
            this.BIWeeklyGb.Name = "BIWeeklyGb";
            this.BIWeeklyGb.Size = new System.Drawing.Size(282, 57);
            this.BIWeeklyGb.TabIndex = 38;
            this.BIWeeklyGb.TabStop = false;
            this.BIWeeklyGb.Text = "BI-Weekly";
            // 
            // BiWeeklyLB
            // 
            this.BiWeeklyLB.CheckOnClick = true;
            this.BiWeeklyLB.FormattingEnabled = true;
            this.BiWeeklyLB.Items.AddRange(new object[] {
            "BI-Weekly ISN Printer - Fry\'s",
            "Run Statement Report"});
            this.BiWeeklyLB.Location = new System.Drawing.Point(5, 15);
            this.BiWeeklyLB.Name = "BiWeeklyLB";
            this.BiWeeklyLB.Size = new System.Drawing.Size(271, 34);
            this.BiWeeklyLB.TabIndex = 13;
            // 
            // Messageslb
            // 
            this.Messageslb.FormattingEnabled = true;
            this.Messageslb.Location = new System.Drawing.Point(12, 41);
            this.Messageslb.Name = "Messageslb";
            this.Messageslb.Size = new System.Drawing.Size(585, 108);
            this.Messageslb.TabIndex = 37;
            // 
            // BTNSGB
            // 
            this.BTNSGB.Controls.Add(this.UnCheckbtn);
            this.BTNSGB.Controls.Add(this.CheckAll);
            this.BTNSGB.Location = new System.Drawing.Point(213, 586);
            this.BTNSGB.Name = "BTNSGB";
            this.BTNSGB.Size = new System.Drawing.Size(384, 64);
            this.BTNSGB.TabIndex = 36;
            this.BTNSGB.TabStop = false;
            // 
            // UnCheckbtn
            // 
            this.UnCheckbtn.Location = new System.Drawing.Point(206, 25);
            this.UnCheckbtn.Name = "UnCheckbtn";
            this.UnCheckbtn.Size = new System.Drawing.Size(75, 23);
            this.UnCheckbtn.TabIndex = 8;
            this.UnCheckbtn.Text = "Clear";
            this.UnCheckbtn.UseVisualStyleBackColor = true;
            this.UnCheckbtn.Click += new System.EventHandler(this.UnCheckbtn_Click);
            // 
            // CheckAll
            // 
            this.CheckAll.Location = new System.Drawing.Point(104, 25);
            this.CheckAll.Name = "CheckAll";
            this.CheckAll.Size = new System.Drawing.Size(75, 23);
            this.CheckAll.TabIndex = 7;
            this.CheckAll.Text = "Select All";
            this.CheckAll.UseVisualStyleBackColor = true;
            this.CheckAll.Click += new System.EventHandler(this.CheckAll_Click);
            // 
            // TimeShowlbl
            // 
            this.TimeShowlbl.AutoSize = true;
            this.TimeShowlbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TimeShowlbl.Location = new System.Drawing.Point(351, 9);
            this.TimeShowlbl.Name = "TimeShowlbl";
            this.TimeShowlbl.Size = new System.Drawing.Size(96, 25);
            this.TimeShowlbl.TabIndex = 32;
            this.TimeShowlbl.Text = "00:00:00";
            // 
            // TimeTXTlbl
            // 
            this.TimeTXTlbl.AutoSize = true;
            this.TimeTXTlbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TimeTXTlbl.Location = new System.Drawing.Point(280, 9);
            this.TimeTXTlbl.Name = "TimeTXTlbl";
            this.TimeTXTlbl.Size = new System.Drawing.Size(65, 25);
            this.TimeTXTlbl.TabIndex = 31;
            this.TimeTXTlbl.Text = "Time:";
            // 
            // DateShowlbl
            // 
            this.DateShowlbl.AutoSize = true;
            this.DateShowlbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DateShowlbl.Location = new System.Drawing.Point(81, 9);
            this.DateShowlbl.Name = "DateShowlbl";
            this.DateShowlbl.Size = new System.Drawing.Size(126, 25);
            this.DateShowlbl.TabIndex = 30;
            this.DateShowlbl.Text = "mm/dd/yyyy";
            // 
            // DateTXTlbl
            // 
            this.DateTXTlbl.AutoSize = true;
            this.DateTXTlbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DateTXTlbl.Location = new System.Drawing.Point(12, 9);
            this.DateTXTlbl.Name = "DateTXTlbl";
            this.DateTXTlbl.Size = new System.Drawing.Size(63, 25);
            this.DateTXTlbl.TabIndex = 29;
            this.DateTXTlbl.Text = "Date:";
            // 
            // Run_button
            // 
            this.Run_button.Location = new System.Drawing.Point(59, 25);
            this.Run_button.Name = "Run_button";
            this.Run_button.Size = new System.Drawing.Size(75, 23);
            this.Run_button.TabIndex = 4;
            this.Run_button.Text = "Run";
            this.Run_button.UseVisualStyleBackColor = true;
            this.Run_button.Click += new System.EventHandler(this.Run_button_Click);
            // 
            // WeekListBox
            // 
            this.WeekListBox.CheckOnClick = true;
            this.WeekListBox.ColumnWidth = 284;
            this.WeekListBox.FormattingEnabled = true;
            this.WeekListBox.Items.AddRange(new object[] {
            "Lock Admin (1)",
            "Pre Billing Report (2)",
            "Invoice Generator (3)",
            "Post Billing Report (4)",
            "Post Bill GL Report (5)",
            "Batch Report- Only Payment Batches (6)",
            "Batch Report- Everything (7)",
            "Batch Report- Detail (8)",
            "NL- Bad Debt Payments (9)",
            "Credit Card Payments (10)",
            "GL Report - 112990 (11)",
            "GL Report - 112999 (12)",
            "Memo Bills- ISN Printer (13)",
            "Memo Bills- FTP (14)"});
            this.WeekListBox.Location = new System.Drawing.Point(6, 19);
            this.WeekListBox.MultiColumn = true;
            this.WeekListBox.Name = "WeekListBox";
            this.WeekListBox.Size = new System.Drawing.Size(573, 109);
            this.WeekListBox.TabIndex = 5;
            // 
            // RunbtnGB
            // 
            this.RunbtnGB.Controls.Add(this.Run_button);
            this.RunbtnGB.Location = new System.Drawing.Point(12, 586);
            this.RunbtnGB.Name = "RunbtnGB";
            this.RunbtnGB.Size = new System.Drawing.Size(195, 64);
            this.RunbtnGB.TabIndex = 35;
            this.RunbtnGB.TabStop = false;
            // 
            // ProcessesGB
            // 
            this.ProcessesGB.Controls.Add(this.WeekListBox);
            this.ProcessesGB.Location = new System.Drawing.Point(12, 155);
            this.ProcessesGB.Name = "ProcessesGB";
            this.ProcessesGB.Size = new System.Drawing.Size(585, 138);
            this.ProcessesGB.TabIndex = 34;
            this.ProcessesGB.TabStop = false;
            this.ProcessesGB.Text = "Daily Processes";
            // 
            // chkBalUtil
            // 
            this.chkBalUtil.AutoSize = true;
            this.chkBalUtil.Location = new System.Drawing.Point(191, 564);
            this.chkBalUtil.Name = "chkBalUtil";
            this.chkBalUtil.Size = new System.Drawing.Size(124, 17);
            this.chkBalUtil.TabIndex = 52;
            this.chkBalUtil.Text = "Stop Balance Utility?";
            this.chkBalUtil.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(609, 662);
            this.Controls.Add(this.chkBalUtil);
            this.Controls.Add(this.MidmonthGB);
            this.Controls.Add(this.AgingUtilchk);
            this.Controls.Add(this.SplitRunCHK);
            this.Controls.Add(this.Closebtn);
            this.Controls.Add(this.ROToolchk);
            this.Controls.Add(this.SundaysISNCB);
            this.Controls.Add(this.BIWeeklyISNCB);
            this.Controls.Add(this.SaturdaysGB);
            this.Controls.Add(this.EOMGB);
            this.Controls.Add(this.EOMISNCB);
            this.Controls.Add(this.MidMonthISNCB);
            this.Controls.Add(this.RMannuallyCB);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BIWeeklyGb);
            this.Controls.Add(this.Messageslb);
            this.Controls.Add(this.BTNSGB);
            this.Controls.Add(this.TimeShowlbl);
            this.Controls.Add(this.TimeTXTlbl);
            this.Controls.Add(this.DateShowlbl);
            this.Controls.Add(this.DateTXTlbl);
            this.Controls.Add(this.RunbtnGB);
            this.Controls.Add(this.ProcessesGB);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(625, 700);
            this.MinimumSize = new System.Drawing.Size(625, 700);
            this.Name = "MainForm";
            this.Text = "NARBOT 1.0";
            this.Resize += new System.EventHandler(this.Window_Resize);
            this.MidmonthGB.ResumeLayout(false);
            this.SaturdaysGB.ResumeLayout(false);
            this.EOMGB.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.BIWeeklyGb.ResumeLayout(false);
            this.BTNSGB.ResumeLayout(false);
            this.RunbtnGB.ResumeLayout(false);
            this.ProcessesGB.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void Window_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                ShowInTaskbar = false;
                BOT_ICON.Visible = true;
            }
            else
            {
                ShowInTaskbar = true;
                BOT_ICON.Visible = false;
            }
        }

        private void BOT_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            WindowState = FormWindowState.Normal;
            Activate();
        }

        private const int CP_NOCLOSE_BUTTON = 0x200;
        private GroupBox MidmonthGB;
        private CheckedListBox MidMonthLB;
        private CheckBox AgingUtilchk;
        private CheckBox SplitRunCHK;
        private Button Closebtn;
        private NotifyIcon BOT_ICON;
        private CheckBox ROToolchk;
        private CheckBox SundaysISNCB;
        private CheckBox BIWeeklyISNCB;
        private GroupBox SaturdaysGB;
        private CheckedListBox SaturdaysLB;
        private GroupBox EOMGB;
        private CheckedListBox EOMLB;
        private CheckBox EOMISNCB;
        private CheckBox MidMonthISNCB;
        private CheckBox RMannuallyCB;
        private GroupBox groupBox1;
        private CheckedListBox SundaysLB;
        private GroupBox BIWeeklyGb;
        private CheckedListBox BiWeeklyLB;
        private ListBox Messageslb;
        private Timer BOT_timer;
        private GroupBox BTNSGB;
        private Button UnCheckbtn;
        private Button CheckAll;
        private Label TimeShowlbl;
        private Label TimeTXTlbl;
        private Label DateShowlbl;
        private Label DateTXTlbl;
        private Button Run_button;
        private CheckedListBox WeekListBox;
        private GroupBox RunbtnGB;
        private GroupBox ProcessesGB;
        private CheckBox chkBalUtil;
    
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams myCp = base.CreateParams;
                myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
                return myCp;
            }
        }

        #endregion

    }
}

