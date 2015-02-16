using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace NARBOT
{
    public partial class MainForm : Form
    {
        Reports bot = new Reports();
        WriteLog logfile = new WriteLog();
        Core method = new Core();
        bool weekdayflag = false;
        bool flag = false;
        bool sundayflag = false;
        bool eomflag = false;
        bool AMflag = false;
        bool midmflag = false;
        bool satflag = false;
        bool minutestodisable = false;
        bool splitweek = false;
        TimeSpan rmannuallychecked;
        string dateforreport = "";
        public MainForm()
        {

            InitializeComponent();
            weekdayflag = method.checkWeekday();
            flag = method.CheckBIWeekly();
            CheckAllItems(flag, BiWeeklyLB);
            sundayflag = method.checkSundays();
            CheckAllItems(sundayflag, SundaysLB);
            eomflag = method.checkEOM();
            CheckAllItems(eomflag, EOMLB);
            midmflag = method.checkMidM();
            CheckAllItems(midmflag, MidMonthLB);
            satflag = method.checkSaturdays();
            CheckAllItems(satflag, SaturdaysLB);
            CheckAllItems(true, WeekListBox);
            BOT_timer.Tick += new EventHandler(BOT_timer_Tick);
            BOT_timer.Interval = (1000) * (1);
            BOT_timer.Enabled = true;
            BOT_timer.Start();


        }

        private void BOT_timer_Tick(object sender, EventArgs e)
        {
            TimeSpan inactive;

            inactive = DateTime.Now.TimeOfDay.Subtract(rmannuallychecked);



            if (!RMannuallyCB.Checked)
            {
                weekdayflag = method.checkWeekday();

                flag = method.CheckBIWeekly();

                sundayflag = method.checkSundays();

                eomflag = method.checkEOM();

                AMflag = method.check12AM();

                midmflag = method.checkMidM();

                satflag = method.checkSaturdays();


                if (!minutestodisable)
                {

                    RMannuallyCB.Enabled = true;

                }
                enable();
                CheckGroup();
            }
            else
            {
                if (inactive.Minutes >= Convert.ToInt32(ConfigurationManager.AppSettings["ManualRunInactive"]))
                {

                    RMannuallyCB.Checked = false;


                }

            }
            dateforreport = DateTime.Today.Year + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0');

            DateShowlbl.Text = DateTime.Now.Date.ToShortDateString().ToString();
            TimeShowlbl.Text = DateTime.Now.ToLongTimeString();
            if (weekdayflag)
            {
                MinutestoDisable("TimetoRunWeek");
                if (DateTime.Now.ToLongTimeString() == ConfigurationManager.AppSettings["TimetoRunWeek"])
                {
                    AutomatedRun();

                }
            }
            else if (satflag)
            {
                MinutestoDisable("TimetoRunSat");
                if (DateTime.Now.ToLongTimeString() == ConfigurationManager.AppSettings["TimetoRunSat"])
                {
                    AutomatedRun();

                }

            }
            else if (sundayflag)
            {
                MinutestoDisable("TimetoRunSun");
                if (DateTime.Now.ToLongTimeString() == ConfigurationManager.AppSettings["TimetoRunSun"])
                {

                    AutomatedRun();
                }

            }
            if (eomflag && DateTime.Now.Day == 1 && AMflag)
            {
                MinutestoDisable("TimetoRunISN");

                if (ConfigurationManager.AppSettings["TimetoRunISN"] == DateTime.Now.ToLongTimeString())
                {

                    AutomatedRun();

                }
            }
        }

        private void MinutestoDisable(string key)
        {

            TimeSpan ttr = Convert.ToDateTime(ConfigurationManager.AppSettings[key]).TimeOfDay;
            if (ttr.Subtract(DateTime.Now.TimeOfDay).Hours == 0 && ttr.Subtract(DateTime.Now.TimeOfDay).Minutes < Convert.ToInt32(ConfigurationManager.AppSettings["DisableMinutes"]) && ttr.Subtract(DateTime.Now.TimeOfDay).Minutes > -1)
            {

                RMannuallyCB.Checked = false;
                RMannuallyCB.Enabled = false;
                minutestodisable = true;


            }

        }

        private void AutomatedRun()
        {

            if (eomflag && DateTime.Today.Day == 1 && AMflag)
            {
                CheckGroupAll(false);
            }
            else
            {
                CheckGroup();

            }
            runningorder();
        }

        public void CheckAllItems(bool state, CheckedListBox lb)
        {

            for (int i = 0; i < lb.Items.Count; i++)
            {

                lb.SetItemChecked(i, state);
            }

        }

        private void disable()
        {


            EnableBtn(Run_button, false);
            EnableBtn(CheckAll, false);
            EnableBtn(UnCheckbtn, false);
            EnableChkLB(WeekListBox, false);
            EnableChkLB(SundaysLB, false);
            EnableChkLB(MidMonthLB, false);
            EnableChkLB(EOMLB, false);
            EnableChkLB(SaturdaysLB, false);
            EnableChkLB(BiWeeklyLB, false);
            RMannuallyCB.Enabled = false;
            MidMonthISNCB.Enabled = false;
            EOMISNCB.Enabled = false;
            BIWeeklyISNCB.Enabled = false;
            SundaysISNCB.Enabled = false;
            chkBalUtil.Enabled = false;
            ROToolchk.Enabled = false;
            AgingUtilchk.Enabled = false;
        }


        private void enable()
        {
            EnableBtn(Run_button, true);
            EnableBtn(UnCheckbtn, true);
            EnableBtn(CheckAll, true);
            EnableChkLB(WeekListBox, true);
            EnableChkLB(BiWeeklyLB, flag);
            EnableChkLB(SundaysLB, sundayflag);
            EnableChkLB(MidMonthLB, midmflag);
            EnableChkLB(EOMLB, eomflag);
            EnableChkLB(SaturdaysLB, satflag);
            if (!minutestodisable)
            {
                RMannuallyCB.Enabled = true;
            }
            MidMonthISNCB.Enabled = true;
            EOMISNCB.Enabled = true;
            BIWeeklyISNCB.Enabled = true;
            SundaysISNCB.Enabled = true;
            chkBalUtil.Enabled = true;
            ROToolchk.Enabled = true;
            AgingUtilchk.Enabled = true;
        }

        private void enable(bool flagenable)
        {
            EnableBtn(Run_button, flagenable);
            EnableBtn(UnCheckbtn, flagenable);
            EnableBtn(CheckAll, flagenable);
            EnableChkLB(WeekListBox, flagenable);
            EnableChkLB(BiWeeklyLB, flagenable);
            EnableChkLB(SundaysLB, flagenable);
            EnableChkLB(MidMonthLB, flagenable);
            EnableChkLB(EOMLB, flagenable);
            EnableChkLB(SaturdaysLB, flagenable);
            if (!minutestodisable)
            {
                RMannuallyCB.Enabled = flagenable;
            }
            MidMonthISNCB.Enabled = flagenable;
            EOMISNCB.Enabled = flagenable;
            BIWeeklyISNCB.Enabled = flagenable;
            SundaysISNCB.Enabled = flagenable;
        }

        private void runSaturdays()
        {
            if (RMannuallyCB.Checked)
            {

                satflag = true;

            }


            if (satflag)
            {
                string log = "";
                Messageslb.Items.Clear();
                disable();
                DateTime starttime = DateTime.Now;
                log = "Running Saturday " + SaturdaysLB.CheckedItems.Count + " process(es) Start time:" + starttime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);



                foreach (int indexChecked in SaturdaysLB.CheckedIndices)
                {


                    try
                    {
                        int rptnumber = Convert.ToInt32(SaturdaysLB.GetItemText(indexChecked)) + 1;
                        log = "Running Saturday process #:" + rptnumber.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);
                        switch (indexChecked)
                        {
                            case 0:

                                if (!chkBalUtil.Checked)
                                {
                                    bot.runBalanceUtility();
                                }
                                else
                                {

                                    log = "Saturday Balance Utility Stopped";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }

                                break;

                            case 1:

                                if (!ROToolchk.Checked)
                                {
                                    bot.runOracleReOpenTool();
                                }
                                else
                                {

                                    log = "Oracle Reopen Tool Stopped";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }



                                break;




                        }
                    }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);


                    }

                    SaturdaysLB.SetItemChecked(indexChecked, false);

                    bot.KillRunningProcesses();

                    rmannuallychecked = DateTime.Now.TimeOfDay;
                }

                DateTime endtime = DateTime.Now;
                log = "Saturday Process(es) Completed, End time:" + endtime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);
                log = "Total Run Time:" + endtime.Subtract(starttime);
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                logfile.writefile("*******************************************", ConfigurationManager.AppSettings[@"LogFile"]);

                chkBalUtil.Checked = false;
                ROToolchk.Checked = false;

                Messageslb.Items.Add(log);

                enable();
            }
        }


        private void runSundays()
        {

            if (RMannuallyCB.Checked)
            {

                sundayflag = true;

            }

            if (sundayflag)
            {
                string log = "";
                Messageslb.Items.Clear();
                disable();
                DateTime starttime = DateTime.Now;
                log = "Running Sunday " + SundaysLB.CheckedItems.Count + " process(es) Start time:" + starttime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);

                splitweek = method.SplitWeek(-7);
                SplitRunCHK.Checked = splitweek;

                foreach (int indexChecked in SundaysLB.CheckedIndices)
                {


                    try
                    {
                        int rptnumber = Convert.ToInt32(SundaysLB.GetItemText(indexChecked)) + 1;
                        log = "Running Sunday process #:" + rptnumber.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);
                        switch (indexChecked)
                        {
                            case 0:
                                if (!SundaysISNCB.Checked)
                                {
                                    if (splitweek)
                                    {
                                        bot.runSundaysISNPrinter("A", "Agencies", splitweek, 1);
                                        bot.runSundaysISNPrinter("A", "Agencies", splitweek, 2);

                                        SundaysLB.SetItemChecked(1, false);
                                    }
                                    else
                                    {
                                        bot.runSundaysISNPrinter("A", "Agencies", false, 0);
                                    }
                                }
                                else
                                {
                                    log = "Agencies ISN Printer Postponed";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }
                                break;

                            case 1:
                                if (!splitweek)
                                {
                                    bot.runStatementReportBIWeekly("Weekly Bill", "A", "LedgerAgency_Weekly_" + dateforreport, true);
                                }
                                break;

                            case 2:
                                if (!SundaysISNCB.Checked)
                                {
                                    if (splitweek)
                                    {
                                        bot.runSundaysISNPrinter("L", "Ledger", splitweek, 1);
                                        bot.runSundaysISNPrinter("L", "Ledger", splitweek, 2);
                                        SundaysLB.SetItemChecked(3, false);
                                    }
                                    else
                                    {
                                        bot.runSundaysISNPrinter("L", "Ledger", false, 0);
                                    }
                                }
                                else
                                {
                                    log = "Ledger ISN Printer Postponed";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }
                                break;

                            case 3:
                                if (!splitweek)
                                {
                                    bot.runStatementReportBIWeekly("Weekly Bill", "L", "Ledger_Weekly_" + dateforreport, true);
                                }
                                break;

                            case 4:
                                if (!AgingUtilchk.Checked)
                                {
                                    bot.runAgingUtility();
                                }
                                else
                                {

                                    log = "Aging Utility Stopped";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }

                                break;



                        }
                    }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);


                    }

                    SundaysLB.SetItemChecked(indexChecked, false);

                    bot.KillRunningProcesses();

                    rmannuallychecked = DateTime.Now.TimeOfDay;
                }

                DateTime endtime = DateTime.Now;
                log = "Sunday Process(es) Completed, End time:" + endtime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);
                log = "Total Run Time:" + endtime.Subtract(starttime);
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                logfile.writefile("*******************************************", ConfigurationManager.AppSettings[@"LogFile"]);

                SundaysISNCB.Checked = false;
                AgingUtilchk.Checked = false;

                Messageslb.Items.Add(log);

                enable();

                SplitRunCHK.Checked = false;
            }
        }
        private void runMidM()
        {
            if (RMannuallyCB.Checked)
            {

                midmflag = true;

            }
            if (midmflag)
            {
                string log = "";
                Messageslb.Items.Clear();
                disable();
                DateTime starttime = DateTime.Now;
                log = "Running Mid-Month " + MidMonthLB.CheckedItems.Count + " process(es) Start time:" + starttime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);

                foreach (int indexChecked in MidMonthLB.CheckedIndices)
                {


                    try
                    {
                        int rptnumber = Convert.ToInt32(MidMonthLB.GetItemText(indexChecked)) + 1;
                        log = "Running Mid-Month process #:" + rptnumber.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);
                        switch (indexChecked)
                        {
                            case 0:
                                if (!MidMonthISNCB.Checked)
                                {
                                    bot.ISNPrinterMIDM("O", "O", "Other");
                                }
                                else
                                {

                                    log = "Mid-Month ISN Printer - Other Postponed";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }
                                break;
                            case 1:
                                if (!MidMonthISNCB.Checked)
                                {
                                    bot.ISNPrinterMIDM("L", "L", "Ledger");
                                }
                                else
                                {

                                    log = "Mid-Month ISN Printer - Ledger Postponed";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }
                                break;
                            case 2:
                                if (!MidMonthISNCB.Checked)
                                {
                                    bot.ISNPrinterMIDM("N", "NL", "NonLedger");
                                }
                                else
                                {

                                    log = "Mid-Month ISN Printer - NonLedger Postponed";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }
                                break;

                            case 3:
                                if (!MidMonthISNCB.Checked)
                                {
                                    bot.ISNPrinterMIDM("A", "L", "Agencies");
                                }
                                else
                                {

                                    log = "Mid-Month ISN Printer - Agencies Postponed";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }
                                break;

                        }
                    }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);


                    }

                    MidMonthLB.SetItemChecked(indexChecked, false);

                    bot.KillRunningProcesses();

                    rmannuallychecked = DateTime.Now.TimeOfDay;
                }

                DateTime endtime = DateTime.Now;
                log = "Mid-Month Process(es) Completed, End time:" + endtime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);
                log = "Total Run Time:" + endtime.Subtract(starttime);
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                logfile.writefile("*******************************************", ConfigurationManager.AppSettings[@"LogFile"]);

                MidMonthISNCB.Checked = false;

                Messageslb.Items.Add(log);

                enable();
            }
        }


        private void runEOM()
        {

            int daybeforemidnight = DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month);
            int dayaftermidnight = DateTime.Today.Day;

            if (RMannuallyCB.Checked)
            {

                eomflag = true;
                AMflag = true;
                daybeforemidnight = DateTime.Today.Day;
                dayaftermidnight = 1;
            }

            if (eomflag)
            {
                string log = "";
                Messageslb.Items.Clear();
                disable();
                DateTime starttime = DateTime.Now;
                log = "Running EOM " + EOMLB.CheckedItems.Count + " process(es) Start time:" + starttime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);

                foreach (int indexChecked in EOMLB.CheckedIndices)
                {


                    try
                    {
                        int rptnumber = Convert.ToInt32(EOMLB.GetItemText(indexChecked)) + 1;
                        log = "Running EOM process #:" + rptnumber.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);
                        switch (indexChecked)
                        {
                            //case 0:
                            //    if (eomflag && daybeforemidnight != 1)
                            //    {
                            //        bot.InvoiceGeneratorEOM();
                            //    }
                            //    break;

                            //case 0:
                            //    if (eomflag && daybeforemidnight ==DateTime.Today.Day)
                            //    {
                            //        if (!EOMBalUtilchk.Checked)
                            //        {
                            //            bot.runBalanceUtility();
                            //        }
                            //        else
                            //        {

                            //            log = "EOM Balance Utility Stopped";
                            //            logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                            //            Messageslb.Items.Add(log);

                            //        }

                            //    }
                            //    break;

                            case 0:
                                if (eomflag && dayaftermidnight == 1 && AMflag)
                                {
                                    bot.runFinanceManager();
                                }
                                break;
                            case 1:
                                if (eomflag && dayaftermidnight == 1 && AMflag)
                                {
                                    if (!EOMISNCB.Checked)
                                    {
                                        bot.runEOMISN("O", "O", "Other");
                                    }
                                    else
                                    {

                                        log = "EOM ISN Printer - Other Posponed";
                                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                        Messageslb.Items.Add(log);

                                    }
                                }
                                break;
                            case 2:
                                if (eomflag && dayaftermidnight == 1 && AMflag)
                                {
                                    if (!EOMISNCB.Checked)
                                    {
                                        bot.runEOMISN("A", "L", "Agencies");
                                    }
                                    else
                                    {

                                        log = "EOM ISN Printer - Agencies Posponed";
                                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                        Messageslb.Items.Add(log);

                                    }
                                }
                                break;

                            case 3:
                                if (eomflag && dayaftermidnight == 1 && AMflag)
                                {
                                    if (!EOMISNCB.Checked)
                                    {
                                        bot.runEOMISN("L", "L", "Ledger");
                                    }
                                    else
                                    {

                                        log = "EOM ISN Printer - Ledger Posponed";
                                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                        Messageslb.Items.Add(log);

                                    }
                                }
                                break;

                            case 4:
                                if (eomflag && dayaftermidnight == 1 && AMflag)
                                {
                                    if (!EOMISNCB.Checked)
                                    {
                                        bot.runEOMISN("N", "NL", "NonLedger");
                                    }
                                    else
                                    {

                                        log = "EOM ISN Printer - NonLedger Posponed";
                                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                        Messageslb.Items.Add(log);

                                    }
                                }
                                break;



                        }
                    }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);


                    }

                    EOMLB.SetItemChecked(indexChecked, false);

                    bot.KillRunningProcesses();

                    rmannuallychecked = DateTime.Now.TimeOfDay;
                }

                DateTime endtime = DateTime.Now;
                log = "EOM Process(es) Completed, End time:" + endtime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);
                log = "Total Run Time:" + endtime.Subtract(starttime);
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                logfile.writefile("*******************************************", ConfigurationManager.AppSettings[@"LogFile"]);

                //EOMBalUtilchk.Checked = false;
                EOMISNCB.Checked = false;

                Messageslb.Items.Add(log);

                enable();
            }
        }



        private void RunNightProcess()
        {
            string log = "";
            Messageslb.Items.Clear();
            disable();
            DateTime starttime = DateTime.Now;
            log = "Running " + WeekListBox.CheckedItems.Count + " process(es) Start time:" + starttime.TimeOfDay.ToString();
            logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
            Messageslb.Items.Add(log);
            int lastmonthday = DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month);

            foreach (int indexChecked in WeekListBox.CheckedIndices)
            {


                try
                {
                    int rptnumber = Convert.ToInt32(WeekListBox.GetItemText(indexChecked)) + 1;
                    log = "Running process #:" + rptnumber.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    Messageslb.Items.Add(log);
                    switch (indexChecked)
                    {


                        case 0:
                            bot.LockAdmin();
                            break;

                        case 1:
                            bot.selectreport("PreBilling", "Pre Billing Report");
                            break;

                        case 2:
                            if (RMannuallyCB.Checked)
                            {

                                if (EOMLB.SelectedIndex == 0)
                                {

                                    eomflag = true;

                                }


                            }

                            if (eomflag && lastmonthday == DateTime.Today.Day)
                            {
                                bot.InvoiceGenerator();
                                bot.InvoiceGeneratorEOM();
                                //EOMLB.SetItemChecked(0, false);


                            }
                            else
                            {
                                bot.InvoiceGenerator();

                            }
                            break;

                        case 3:
                            bot.selectreport("PostBilling", "Post Billing Report");
                            break;

                        case 4:
                            bot.selectreport("PostBillGL", "GLResults");
                            break;

                        case 5:
                            bot.selectreport("PaymentBatches", "Batch Report - Payment Batches");
                            break;

                        case 6:
                            bot.selectreport("PaymentEverything", "Batch Report - Everything");
                            break;

                        case 7:
                            bot.selectreport("PaymentDetail", "Batch Report - Detail");
                            break;

                        case 8:
                            bot.selectreport("NLBadDebtPayments", "Non-Ledger Bad Debt Payments");
                            break;

                        case 9:
                            bot.selectreport("CreditCardPayments", "Credit Card Payments");
                            break;

                        case 10:
                            bot.selectreport("GLAditional_112990", "GL Report 112990");
                            break;

                        case 11:
                            bot.selectreport("GLAditional_112999", "GL Report 112999");
                            break;

                        case 12:
                            bot.runISNPrinter();
                            break;

                        case 13:
                            bot.MemoBills();
                            break;
                    }
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    Messageslb.Items.Add(log);


                }

                WeekListBox.SetItemChecked(indexChecked, false);

                bot.KillRunningProcesses();

                rmannuallychecked = DateTime.Now.TimeOfDay;
            }


            DateTime endtime = DateTime.Now;
            log = "Process(es) Completed, End time:" + endtime.TimeOfDay.ToString();
            logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
            Messageslb.Items.Add(log);
            log = "Total Run Time:" + endtime.Subtract(starttime);
            logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
            logfile.writefile("*******************************************", ConfigurationManager.AppSettings[@"LogFile"]);

            Messageslb.Items.Add(log);

            enable();

        }

        private void runBIWeekly()
        {

            if (RMannuallyCB.Checked)
            {

                flag = true;

            }


            if (flag)
            {
                string log = "";
                Messageslb.Items.Clear();
                disable();
                DateTime starttime = DateTime.Now;
                log = "Running BI-Weekly" + BiWeeklyLB.CheckedItems.Count + " process(es) Start time:" + starttime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);

                splitweek = method.SplitWeek(-14);
                SplitRunCHK.Checked = splitweek;

                foreach (int indexChecked in BiWeeklyLB.CheckedIndices)
                {

                    try
                    {
                        switch (indexChecked)
                        {

                            case 0:

                                log = "Running BI-Weekly ISN Printer - Fry's";
                                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                Messageslb.Items.Add(log);
                                if (!BIWeeklyISNCB.Checked)
                                {
                                    if (splitweek)
                                    {
                                        bot.runBIWeeklyISNPrinterFry(splitweek, 1);

                                        bot.runBIWeeklyISNPrinterFry(splitweek, 2);

                                        BiWeeklyLB.SetItemChecked(1, false);
                                    }
                                    else
                                    {
                                        bot.runBIWeeklyISNPrinterFry(false, 0);
                                    }
                                }
                                else
                                {

                                    log = "BI-Weekly ISN Printer - Fry's Postponed";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                }
                                break;

                            case 1:
                                if (!splitweek)
                                {
                                    log = "Running Statement Report";
                                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                                    Messageslb.Items.Add(log);

                                    bot.runStatementReportBIWeekly("*", "L", "Frys_" + dateforreport, false);
                                }
                                break;
                        }

                    }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                        Messageslb.Items.Add(log);


                    }

                    BiWeeklyLB.SetItemChecked(indexChecked, false);

                    bot.KillRunningProcesses();

                    rmannuallychecked = DateTime.Now.TimeOfDay;
                }
                DateTime endtime = DateTime.Now;
                log = "Process(es) Completed, End time:" + endtime.TimeOfDay.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                Messageslb.Items.Add(log);
                log = "Total Run Time:" + endtime.Subtract(starttime);
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                logfile.writefile("*******************************************", ConfigurationManager.AppSettings[@"LogFile"]);

                BIWeeklyISNCB.Checked = false;

                Messageslb.Items.Add(log);

                enable();

                SplitRunCHK.Checked = false;
            }


        }

        private void runningorder()
        {



            if (RMannuallyCB.Checked)
            {
                RunNightProcess();
                runBIWeekly();
                runSundays();
                runMidM();
                runSaturdays();
                runEOM();
                RMannuallyCB.Checked = false;
            }
            else
            {

                RunNightProcess();
                if (flag)
                {
                    runBIWeekly();
                }

                if (sundayflag)
                {

                    runSundays();

                }

                if (midmflag)
                {

                    runMidM();

                }

                if (satflag)
                {

                    runSaturdays();

                }

                if (eomflag)
                {

                    runEOM();

                }
                minutestodisable = false;

            }

            GC.Collect();
            GC.WaitForFullGCComplete();

        }

        private void Run_button_Click(object sender, EventArgs e)
        {
            runningorder();

        }

        private void CheckAll_Click(object sender, EventArgs e)
        {
            if (RMannuallyCB.Checked)
            {

                CheckAllItems(true, WeekListBox);

                CheckAllItems(true, BiWeeklyLB);


                CheckAllItems(true, SundaysLB);


                CheckAllItems(true, EOMLB);

                CheckAllItems(true, SaturdaysLB);

                CheckAllItems(true, MidMonthLB);

            }
            else
            {
                CheckAllItems(true, WeekListBox);
                if (flag)
                {
                    CheckAllItems(true, BiWeeklyLB);
                }
                if (sundayflag)
                {
                    CheckAllItems(true, SundaysLB);
                }
                if (eomflag)
                {
                    CheckAllItems(true, EOMLB);
                }
                if (satflag)
                {
                    CheckAllItems(true, SaturdaysLB);
                }
                if (midmflag)
                {
                    CheckAllItems(true, MidMonthLB);
                }
            }
        }

        private void UnCheckbtn_Click(object sender, EventArgs e)
        {
            if (RMannuallyCB.Checked)
            {

                CheckAllItems(false, WeekListBox);

                CheckAllItems(false, BiWeeklyLB);


                CheckAllItems(false, SundaysLB);


                CheckAllItems(false, EOMLB);

                CheckAllItems(false, SaturdaysLB);

                CheckAllItems(false, MidMonthLB);

            }
            else
            {


                CheckAllItems(false, WeekListBox);
                if (flag)
                {
                    CheckAllItems(false, BiWeeklyLB);
                }
                if (sundayflag)
                {
                    CheckAllItems(false, SundaysLB);
                }
                if (eomflag)
                {
                    CheckAllItems(false, EOMLB);
                }
                if (satflag)
                {
                    CheckAllItems(false, SaturdaysLB);
                }
                if (midmflag)
                {
                    CheckAllItems(false, MidMonthLB);
                }
            }
        }

        private void EnableBtn(Button btn, bool state)
        {

            btn.Enabled = state;
        }

        public void EnableChkLB(CheckedListBox lb, bool state)
        {

            lb.Enabled = state;
        }

        private void RMannuallyCB_CheckedChanged(object sender, EventArgs e)
        {


            if (RMannuallyCB.Checked)
            {
                enable(true);
                CheckGroupAll(true);
                rmannuallychecked = DateTime.Now.TimeOfDay;
            }
            else
            {

                CheckGroup();


            }
        }

        private void CheckGroup()
        {

            CheckAllItems(true, WeekListBox);
            CheckAllItems(flag, BiWeeklyLB);
            CheckAllItems(sundayflag, SundaysLB);
            CheckAllItems(eomflag, EOMLB);
            CheckAllItems(midmflag, MidMonthLB);
            CheckAllItems(satflag, SaturdaysLB);


        }

        private void CheckGroupAll(bool check)
        {
            CheckAllItems(check, WeekListBox);
            CheckAllItems(check, BiWeeklyLB);
            CheckAllItems(check, SundaysLB);

            if (eomflag && DateTime.Today.Day == 1 && AMflag)
            {
                CheckAllItems(true, EOMLB);
            }
            else
            {

                CheckAllItems(check, EOMLB);
            }

            CheckAllItems(check, MidMonthLB);
            CheckAllItems(check, SaturdaysLB);

        }

        private void Closebtn_Click(object sender, EventArgs e)
        {
            bot.KillRunningProcesses();
            GC.Collect();
            GC.WaitForFullGCComplete();
            this.Dispose();
        }

  
    }
}
