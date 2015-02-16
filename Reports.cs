using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TestStack.White;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems;
using TestStack.White.UIItems.ListBoxItems;
using System.Configuration;
using TestStack.White.UIItems.TabItems;
using System.Globalization;
using TestStack.White.UIItems.WindowStripControls;
using TestStack.White.UIItems.TableItems;
using TestStack.White.UIItems.Finders;
using System.Windows.Automation;

namespace NARBOT
{
    class Reports
    {
        private Core method=new Core();
        private ScreenShots screencapture = new ScreenShots();
        private EMail mail=new EMail();
        WriteLog logfile = new WriteLog();
        //************Select the report to run******************

        public void LockAdmin() {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
           method.launchlockadmin();
           visitLockAdmin();
           screencapture.printScreen("Lock Admin");
           mail.email_send("Lock Admin", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Lock Admin.png");
           KillRunningProcesses();
           tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        }

        public void InvoiceGenerator()
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchInvoiceGenerator();
            InvoiceGeneratorSelections();
            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        }
        public void InvoiceGeneratorEOM()
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchInvoiceGenerator();
            InvoiceGeneratorSelectionsEOM();
            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        }
        public void selectreport(string reportname, string partialfilename)
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchOracleReports();
            //string path = method.buildpath(ConfigurationManager.AppSettings[@"SavePath"]);
            //string filename = method.buildfilename(partialfilename);
            string reportpath=ConfigurationManager.AppSettings[@"ReportPath"]+@"\" + DateTime.Today.Month.ToString() + DateTime.Today.Day.ToString() + DateTime.Today.Year.ToString()+@"\";
            //*****************************************************************
            if (reportname == "PreBilling") {

                runBillingReport(method.mainwind, "Not Invoiced/Billed", partialfilename,reportpath );
            
            }
            if (reportname == "PostBilling")
            {

                runBillingReport(method.mainwind, "Invoiced/Billable", partialfilename, reportpath);

            }

            if (reportname == "PostBillGL")
            {

                runGL(method.mainwind, "All", ConfigurationManager.AppSettings[@"QlikviewPath"], partialfilename, "Complete");

            }

            if(reportname=="PaymentBatches"){

                runPayments(method.mainwind, reportpath, partialfilename, "Only Payment  Batches");
            
            
            }

            if (reportname == "PaymentEverything")
            {

                runPayments(method.mainwind, reportpath, partialfilename, "");


            }

            if (reportname == "PaymentDetail")
            {

                runPayments(method.mainwind, reportpath, partialfilename, "Detail");


            }
            if (reportname == "NLBadDebtPayments")
            {

                runPaymentsCustomer(method.mainwind, "NL", reportpath, partialfilename, "All Payments(Payment Method/Customer Type)");


            }

            if (reportname == "CreditCardPayments")
            {

                runPaymentsCustomer(method.mainwind, "", reportpath, partialfilename, "Credit Card Only(By Card Type)");


            }

            if (reportname == "GLAditional_112990")
            {

                runGLAditional(method.mainwind,"All",reportpath,partialfilename,"112990");


            }

            if (reportname == "GLAditional_112999")
            {

                runGLAditional(method.mainwind, "All", reportpath, partialfilename, "112999");


            }
            //*************************************************************
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }

        }

        //************* Processes to Run **************************

        private void runBillingReport(Window wind, string type, string name, string reportpath) {
           
            method.gerptname(wind, 18);
            method.getButton(wind, "Select");
            Window report = method.getWindow("Billing Report");
            method.getRadioButton("Publishable", report, 0);
            method.getRadioButton(type, report, 0);
            method.getRadioButton("Billing", report, 0);
            method.getRadioButton("Orders Marked to Bill", report, 0);
            method.getRadioButton("Don't Care", report, 1);
            method.getRadioButton("Include", report, 0);
            method.getRadioButton("Don't Care", report, 2);
            method.getRadioButton("Include", report, 1);
            method.SetDateTodayTomorrow(0, report, 0);
            method.SetDateTodayTomorrow(1, report, 1);
            screencapture.printScreen(name);
            method.getButton(report, "View");
            method.printpreviewwindow(reportpath, name);

            mail.email_send(name, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + name + ".png");

            method.closewindow(report);
            method.closewindow(method.mainwind);

            KillRunningProcesses();
           
        }

        private void runGLAditional(Window wind, string limitTo, string path, string filename, string account)
        {
            
            method.gerptname(wind, 20);
            method.getButton(wind, "Select");
            Window report = method.getWindow("GL Reports");
            method.getRadioButton("By GL Account by Transaction Type", report, 0);
            method.getCheckBox("Show Detail", report, 0);
            method.getRadioButton("Complete", report, 0);
            method.getRadioButton(limitTo, report, 0);
            method.getRadioButton("Approved", report, 0);
            method.getRadioButton("Don't Care - All", report, 0);
            ComboBox limittoacct = method.getComboBoxByClass("TComboBox", report, 1);
            limittoacct.Select(account);
            method.getRadioButton("Use Filter for Both", report, 0);
            method.SetBalFirstDates(0, report, 0);
            method.SetDateTodayTomorrow(1, report, 1);

            screencapture.printScreen(filename);

            method.getButton(report, "View");
            method.app.WaitWhileBusy();
            method.printpreviewwindow(path, filename);

            mail.email_send(filename, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + filename + ".png");

            method.closewindow(report);
            method.closewindow(method.mainwind);

            KillRunningProcesses();
            
        }

        private void runGL(Window wind, string limitTo, string path, string filename, string option)
        {
          
            method.gerptname(wind, 20);
            method.getButton(wind, "Select");
            Window report = method.getWindow("GL Reports");
            method.getRadioButton(option, report, 0);
            method.getRadioButton(limitTo, report, 0);
            method.getRadioButton("Approved", report, 0);
            method.getRadioButton("Don't Care - All", report, 0);
            if (option == "Complete")
            {

                method.getCheckBox("Save To File", report, 0);

            }
            method.getRadioButton("Use Filter for Both", report, 0);
            method.SetBalFirstDates(0, report, 0);
            method.SetDateTodayTomorrow(1, report, 1);

            screencapture.printScreen(filename);

            method.getButton(report, "View");
            if (option == "Complete")
            {
                method.createPath(path);
                Window saveas = method.getWindow("Save As");
                KeyStrokes sendtext = new KeyStrokes();
                sendtext.sendvalue(path + @"\" + filename + ".csv");
                method.getButton(saveas, "Save");

            }
            method.app.WaitWhileBusy();

            mail.email_send(filename, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + filename + ".png");

            method.closewindow(report);
            method.closewindow(method.mainwind);

            KillRunningProcesses();
          
        }


        private void InvoiceGeneratorSelections() {
            Window invgen = method.mainwind;
            method.getCheckBox("Store Invoices", invgen, 0);
            method.getCheckBox("Include Orders", invgen, 0);
            method.getCheckBox("Include Debits", invgen, 0);
            method.SetDateTodayTomorrowCHK(0, invgen, 0);
            method.SetDateTodayTomorrowCHK(1, invgen, 1);
            method.SetDateTodayTomorrow(2, invgen, 0);
            method.SetDateTodayTomorrow(3, invgen, 0);
            method.getButton(invgen, "Generate");
            method.iddleStatePorgressBar(invgen);
            method.app.WaitWhileBusy();
            method.iddleState("InvoiceGenerator");
            screencapture.printScreen("Invoice Generator");

            mail.email_send("Invoice Generator", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Invoice Generator" + ".png");

            method.getButton(invgen, "View Log");
            method.app.WaitWhileBusy();

            method.attachnotepad();

            KillRunningProcesses();

            
        }

        private void InvoiceGeneratorSelectionsEOM()
        {
            Window invgen = method.mainwind;
            method.getCheckBox("Store Invoices", invgen, 0);
            method.getCheckBox("Include Orders", invgen, 0);
            method.getCheckBox("Include Debits", invgen, 0);
            method.SetBalFirstDatesCHK(0, invgen, 0);
            method.SetBalFirstDatesCHK(1, invgen, 1);
            method.SetDateTodayTomorrow(2, invgen, 0);
            method.SetDateTodayTomorrow(3, invgen, 0);
            method.getButton(invgen, "Generate");
            method.iddleStatePorgressBar(invgen);
            method.app.WaitWhileBusy();
            method.iddleState("InvoiceGenerator");
            screencapture.printScreen("Invoice Generator EOM");

            mail.email_send("Invoice Generator EOM", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Invoice Generator EOM" + ".png");

            method.getButton(invgen, "View Log");
            method.app.WaitWhileBusy();

            method.attachnotepad();
            KillRunningProcesses();
            
        }

        private void runPayments(Window wind, string reportpath, string filename, string Type)
        {
           
            method.gerptname(wind, 8);
            method.getButton(wind, "Select");
            Window report = method.getWindow("Payment Selection Report Form");
            method.getRadioButton("Batch Report", report, 0);
            if(Type!=""){
            method.getCheckBox(Type, report, 0);
            }
            method.getRadioButton("Approved", report, 1);
            method.SetDateTodayTomorrow(0, report, 0);
            method.SetDateTodayTomorrow(1, report, 1);

            screencapture.printScreen(filename);

            method.getButton(report, "View");
            method.app.WaitWhileBusy();
            method.printpreviewwindow(reportpath,filename);

            mail.email_send(filename, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + filename + ".png");

            method.closewindow(report);
            method.closewindow(method.mainwind);
            KillRunningProcesses();
           

        }

        private void runPaymentsCustomer(Window wind, string CtoRun, string path, string filename, string Type)
        {
            
            method.gerptname(wind, 8);
            method.getButton(wind, "Select");
            Window report = method.getWindow("Payment Selection Report Form");
            method.getRadioButton(Type, report, 0);
            method.getRadioButton("All Credits/Debits", report, 0);
            method.getRadioButton("Payment Effective Date", report, 0);
            method.getRadioButton("Don't Care - All Payments", report, 0);
            method.getRadioButton("Show Only Active", report, 0);
            if (Type == "Credit Card Only(By Card Type)") {

                method.getRadioButton("Don't Care", report, 0);
            
            }
            if (Type == "Credit Card Only(By Card Type)")
            {

                method.getRadioButton("Don't Care - All Payments", report, 1);

            }
            else {

                method.getRadioButton("Only Bad Debt Payments", report, 0);
            
            }
           
            method.getRadioButton("Approved", report, 1);
            if (Type == "All Payments(Payment Method/Customer Type)")
            {
                
                method.eraseCustomerType(report, "", 0, 195, 169);
                method.selectcustomerforrep(CtoRun);
                Window custwind = method.getWindow("Select Customer Types For Report");
                method.getButton(custwind, "Select");
                method.SetDateTodayTomorrow(0, report, 0);
                method.SetDateTodayTomorrow(1, report, 1);
            }
            else 
            {

                method.SetDateTodayTomorrow(0, report, -1);
                method.SetDateTodayTomorrow(1, report, 0);
            
            }

            screencapture.printScreen(filename);

            method.getButton(report, "View");
            method.app.WaitWhileBusy();
            method.printpreviewwindow(path, filename);

            mail.email_send(filename, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + filename + ".png");

            method.closewindow(report);
            method.closewindow(method.mainwind);
            KillRunningProcesses();
           
        }


        public void runISNPrinter() 
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchISNPrinter();
            method.getButton(method.mainwind, "Generate Invoices");
            Window invprinter = method.getWindow("Invoice Printer");
            ComboBox company = method.getComboBoxByClass("TComboBox", invprinter, 0);
            company.Select("Oregonian");
            method.getCheckBox("Clear Table Data if Company Selected?", invprinter, 0);
            method.getButton(invprinter, "Limit to Customer Type");
            method.selectcustomerforrepISN("L");
            method.selectcustomerforrepISN("NL-B,NL-G");
            Window custwind = method.getWindow("Customer Type Selection");
            method.getButton(custwind, "OK");
            method.getCheckBox("", invprinter, 0);
            ComboBox balancedue = method.getComboBoxByClass("TComboBox", invprinter, 2);
            balancedue.Select("All");
            method.SetDateTodayTomorrow(2, invprinter, 0);
            method.SetDateTodayTomorrow(3, invprinter, 1);
            method.getCheckBox("Include Printed Invoices?", invprinter, 0);
            method.getUnCheckBox("Process Payments and Credits?", invprinter, 0);
            method.getCheckBox("Create Separate Typography Charge Details?", invprinter, 0);
            method.getRadioButton("Actual Run Date Cost", invprinter, 0);
            method.getCheckBox("Separate Discounts?", invprinter, 0);
            method.getCheckBox("Separate Premiums?", invprinter, 0);
            method.getCheckBox("Separate Color Charges?", invprinter, 0);
            method.SetDateTodayTomorrow(4, invprinter, 0);
            TextBox billingperiod = method.getTextBox("", invprinter, 4);
            billingperiod.BulkText = DateTime.Today.ToShortDateString() + " - " + DateTime.Today.ToShortDateString();
            method.getButton(invprinter, "Create Data");
            method.app.WaitWhileBusy();
            try{   
            if (invprinter.HasPopup())
            {
                Window warning = invprinter.ModalWindow("Warning");
                method.getButton(warning, "OK");
            }

            if (invprinter.HasPopup())
            {
                Window warning = invprinter.ModalWindow("Error");
                method.getButton(warning, "OK");
            }

            if (invprinter.HasPopup())
            {
                Window warning = invprinter.ModalWindow("Error");
                method.getButton(warning, "OK");
            }

         
              }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                      

                    }

            method.iddleStatePorgressBar(invprinter);
            method.iddleState("OracleISNPrinter");
            method.app.WaitWhileBusy();
            screencapture.printScreen("Memo Bills- ISNPrinter");

            mail.email_send("Memo Bills- ISNPrinter", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Memo Bills- ISNPrinter" + ".png");
            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
           
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }

           
           
        }
        public void MemoBills() {
           
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
                    string fname = method.buildfilename("MemoInv_");
                    CrystalReports memoinvrep = new CrystalReports();
                    memoinvrep.CRMemoInv(fname);

                    PDFtoXML conversion = new PDFtoXML();
                    conversion.ConvertPDFtoXML(memoinvrep.memoinvpdfreportandpath);

                    FTPSend ftp = new FTPSend();

                    ftp.ftpcopy(memoinvrep.memoinvpdfreportandpath, fname + ".pdf");
                    ftp.ftpcopy(memoinvrep.memoinvpdfreportandpath.Replace(".pdf", ".xml"), fname + ".xml");
                   
                     mail.email_send("Memo Invoices to Presteligence", "");

                     KillRunningProcesses();

          
                    tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        }
       
        public void runBIWeeklyISNPrinterFry(bool split, int splitnumber)
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchISNPrinter();
            method.getButton(method.mainwind, "Generate Statements");
            Window invprinter = method.getWindow("Statement Printer");
            if (Convert.ToBoolean(ConfigurationManager.AppSettings["Maximize"]))
            {
                method.getButton(invprinter, "Maximize");
            }
            ComboBox company = method.getComboBoxByClass("TComboBox", invprinter, 0);
            company.Select("Oregonian");
            TextBox statementkey = method.getTextBox("", invprinter, 0);
            statementkey.BulkText = "L";
            method.getCheckBox("Clear Table Data if Company Selected?", invprinter, 0);
            ComboBox disttype = method.getComboBoxByClass("TComboBox", invprinter, 2);
            disttype.Select("FrysBiWeekly");
            method.dotabs(1);
            KeyStrokes key = new KeyStrokes();
            key.sendvalue("2000122521");
            method.getRadioButton("Exclude Agencies", invprinter, 0);
            method.getRadioButton("All Customers", invprinter, 1);
            ComboBox frequency = method.getComboBoxByClass("TComboBox", invprinter, 3);
            frequency.Items.Clear();
            ComboBox period = method.getComboBoxByClass("TComboBox", invprinter, 4);

            if (split)
            {
                if (splitnumber == 1)
                {
                    period.Select(PredefinedPeriodSplit());
                    method.SetBalFirstDatesSplit(1, invprinter, -1);
                    method.SetBalFirstDatesSplit(2, invprinter, 0);
                }
                else
                {

                    period.Select(PredefinedPeriod());
                    method.SetBalFirstDates(1, invprinter, 0);
                    method.SetBalFirstDates(2, invprinter, 1);
                }
            }
            else
            {
                period.Select(PredefinedPeriod());
                method.SetBalFirstDates(1, invprinter, 0);
                method.SetBalFirstDates(2, invprinter, 1);
            }

           
            method.getCheckBox("Use Different Activity Date Range?", invprinter, 0);
            method.getCheckBox("Age using Differrent Acivity End Date?", invprinter, 0);
            DateTime prevfriday=method.previousFriday();

            if (split)
            {
                string date = "";
                DateTime fdate;
                if (splitnumber == 2)
                {


                    date = DateTime.Now.Month.ToString() + "/1/" + DateTime.Now.Year.ToString();



                    fdate = Convert.ToDateTime(date);

                    method.SetDayOfWeek(3, invprinter, fdate);

                    method.SetDayOfWeek(4, invprinter, DateTime.Today.AddDays(1));
                }
                else
                {
                    date = DateTime.Now.Month.ToString() + "/1/" + DateTime.Now.Year.ToString();
                    fdate = Convert.ToDateTime(date);
                    method.SetDayOfWeek(3, invprinter, prevfriday);
                    method.SetDayOfWeek(4, invprinter, fdate);

                }

            }
            else
            {
                method.SetDayOfWeek(3, invprinter, prevfriday);
                method.SetDayOfWeek(4, invprinter, DateTime.Today.AddDays(1));
            }

          

           method.getUnCheckBox("Include Closed Transactions?", invprinter, 0);
           method.getCheckBox("Include Printed Invoices?", invprinter, 0);
           method.getRadioButton("Balance Forward", invprinter, 0);
           method.getCheckBox("Detail Line for Each Run Date?", invprinter, 0);
           method.getRadioButton("Actual Run Date Cost", invprinter, 0);
           method.getCheckBox("Separate Color Charges?", invprinter, 0);
           method.getCheckBox("Separate Position Premium?", invprinter, 0);
           method.getCheckBox("Separate Material Charges?", invprinter, 0);
           method.getCheckBox("Separate Discounts?", invprinter, 0);
           method.getCheckBox("Create Single Contract Discount Detail for Each Invoice?", invprinter, 0);
           method.getCheckBox("Details Typo Cost?", invprinter, 0);
           method.getCheckBox("Details Color Cost?", invprinter, 0);
           method.getCheckBox("Create Short Rate / Rebate Details?", invprinter, 0);
           method.getCheckBox("Separate Typography Details?", invprinter, 0);
           method.getRadioButton("No Statement Number", invprinter, 0);
           method.getRadioButton("Customer Name", invprinter, 0);
           method.getUnCheckBox("Create Contract Data for Subrep?", invprinter, 0);
           method.SetDateTodayTomorrow(6, invprinter, 0);
           method.dotabs(1);

           if (split)
           {
               if (splitnumber == 1)
               {
                   key.sendvalue(prevfriday.ToShortDateString().ToString() + " - " + prevfriday.Month.ToString() + "/" + method.numberdays(prevfriday).ToString() + "/" + prevfriday.Year.ToString());

               }
               else
               {
                   key.sendvalue(DateTime.Today.Month.ToString() + "/1/" + DateTime.Today.Year.ToString() + " - " + DateTime.Today.ToShortDateString());

               }



           }
           else
           {
               key.sendvalue(prevfriday.ToShortDateString() + " - " + DateTime.Today.ToShortDateString());
           }



           method.getButton(invprinter, "Create Data");
           method.app.WaitWhileBusy();

           try
           {
               if (invprinter.HasPopup())
               {
                   Window warning = invprinter.ModalWindow("Warning");
                   method.getButton(warning, "OK");
               }

           }
           catch (Exception ex)
           {

               log = ex.Message.ToString();
               logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);


           }

           method.iddleStatePorgressBar(invprinter);

           method.iddleState("OracleISNPrinter");
           method.app.WaitWhileBusy();
           if (split)
           {
               screencapture.printScreen("BI-Weekly ISN Printer " + splitnumber.ToString() + " - Fry's");

               mail.email_send("BI-Weekly ISN Printer " + splitnumber.ToString() + " - Fry's", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "BI-Weekly ISN Printer " + splitnumber.ToString() + " - Fry's" + ".png");

           }
           else
           {
               screencapture.printScreen("BI-Weekly ISN Printer - Fry's");

               mail.email_send("BI-Weekly ISN Printer - Fry's", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "BI-Weekly ISN Printer - Fry's" + ".png");
           }

           if (split)
           {

               if (splitnumber == 1)
               {

                   runStatementReportBIWeekly("*", "L", "Frys_" + prevfriday.Year.ToString() + prevfriday.Month.ToString().PadLeft(2, '0') + method.numberdays(prevfriday).ToString(), false);

               }
               else
               {

                   runStatementReportBIWeekly("*", "L", "Frys_" + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0'), false);
               }

           }




           KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }

        }


        


        public void runSundaysISNPrinter(string statkey,string type, bool split, int splitnumber)
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchISNPrinter();
            method.getButton(method.mainwind, "Generate Statements");
            Window invprinter = method.getWindow("Statement Printer");
            if (Convert.ToBoolean(ConfigurationManager.AppSettings["Maximize"]))
            {
                method.getButton(invprinter, "Maximize");
            }
            ComboBox company = method.getComboBoxByClass("TComboBox", invprinter, 0);
            company.Select("Oregonian");
            TextBox statementkey = method.getTextBox("", invprinter, 0);
            statementkey.BulkText = statkey;
            method.getCheckBox("Clear Table Data if Company Selected?", invprinter, 0);
            ComboBox disttype = method.getComboBoxByClass("TComboBox", invprinter, 2);
            disttype.Select("Weekly Bill");
            if (type == "Ledger" )
            {
                method.getRadioButton("Exclude Agencies", invprinter, 0);
            }
           
            if (type == "Agencies")
            {
                method.getRadioButton("Agencies Only Aging By Client", invprinter, 0);
            }
            method.getRadioButton("All Customers", invprinter, 1);
            ComboBox frequency = method.getComboBoxByClass("TComboBox", invprinter, 3);
            frequency.Items.Clear();
            ComboBox period = method.getComboBoxByClass("TComboBox", invprinter, 4);
            if (split)
            {
                if (splitnumber == 1)
                {
                    period.Select(PredefinedPeriodSplit());
                    method.SetBalFirstDatesSplit(1, invprinter, -1);
                    method.SetBalFirstDatesSplit(2, invprinter, 0);
                }
                else {

                    period.Select(PredefinedPeriod());
                    method.SetBalFirstDates(1, invprinter, 0);
                    method.SetBalFirstDates(2, invprinter, 1);
                }
            }
            else
            {
                period.Select(PredefinedPeriod());
                method.SetBalFirstDates(1, invprinter, 0);
                method.SetBalFirstDates(2, invprinter, 1);
            }
            method.getCheckBox("Use Different Activity Date Range?", invprinter, 0);
            method.getCheckBox("Age using Differrent Acivity End Date?", invprinter, 0);
            if (split)
            {
                string date="";
                DateTime fdate;
                if (splitnumber == 2)
                {


                    date = DateTime.Now.Month.ToString() + "/1/" + DateTime.Now.Year.ToString();

              

                   fdate = Convert.ToDateTime(date);

                    method.SetDayOfWeek(3, invprinter, fdate);

                    method.SetDayOfWeek(4, invprinter, DateTime.Today.AddDays(1));
                }
                else {
                    date = DateTime.Now.Month.ToString() + "/1/" + DateTime.Now.Year.ToString();
                    fdate = Convert.ToDateTime(date);
                    method.SetDayOfWeek(3, invprinter, DateTime.Today.AddDays(-6));
                    method.SetDayOfWeek(4, invprinter, fdate);
                
                }

            }
            else
            {
                method.SetDayOfWeek(3, invprinter, DateTime.Today.AddDays(-6));
                method.SetDayOfWeek(4, invprinter, DateTime.Today.AddDays(1));
            }
            method.getUnCheckBox("Include Closed Transactions?", invprinter, 0);
            method.getCheckBox("Include Printed Invoices?", invprinter, 0);
            method.getRadioButton("Balance Forward", invprinter, 0);
            method.getCheckBox("Detail Line for Each Run Date?", invprinter, 0);
            method.getRadioButton("Actual Run Date Cost", invprinter, 0);
            method.getCheckBox("Separate Color Charges?", invprinter, 0);
            method.getCheckBox("Separate Position Premium?", invprinter, 0);
            method.getCheckBox("Separate Material Charges?", invprinter, 0);
            method.getCheckBox("Separate Discounts?", invprinter, 0);
            method.getCheckBox("Create Single Contract Discount Detail for Each Invoice?", invprinter, 0);
            method.getCheckBox("Details Typo Cost?", invprinter, 0);
            method.getCheckBox("Details Color Cost?", invprinter, 0);
            method.getCheckBox("Create Short Rate / Rebate Details?", invprinter, 0);
            method.getCheckBox("Separate Typography Details?", invprinter, 0);
            method.getRadioButton("No Statement Number", invprinter, 0);
            method.getRadioButton("Customer Name", invprinter, 0);
            method.getUnCheckBox("Create Contract Data for Subrep?", invprinter, 0);
            method.SetDateTodayTomorrow(6, invprinter, 0);
            method.dotabs(1);
            KeyStrokes key = new KeyStrokes();
            if (split)
            {
                if (splitnumber == 1)
                {
                    key.sendvalue(DateTime.Today.AddDays(-6).ToShortDateString() + " - " + DateTime.Today.AddMonths(-1).Month.ToString() + "/" + method.numberdays(DateTime.Today.AddMonths(-1)).ToString() + "/" + DateTime.Today.AddMonths(-1).Year.ToString());

                }
                else {

                    key.sendvalue(DateTime.Today.Month.ToString() + "/1/" + DateTime.Today.Year.ToString() + " - " + DateTime.Today.ToShortDateString());
                }



            }
            else
            {
                key.sendvalue(DateTime.Today.AddDays(-6).ToShortDateString() + " - " + DateTime.Today.ToShortDateString());
            }

            
            method.getButton(invprinter, "Create Data");
            method.app.WaitWhileBusy();

            try{
            if (invprinter.HasPopup())
            {
                Window warning = invprinter.ModalWindow("Warning");
                method.getButton(warning, "OK");
            }
            }
            catch (Exception ex)
            {

                log = ex.Message.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);


            }
            method.iddleStatePorgressBar(invprinter);

            method.app.WaitWhileBusy();

            if (split)
            {
                screencapture.printScreen("Sundays ISN Printer Split " + splitnumber.ToString() + " - " + type);
                mail.email_send("Sundays ISN Printer Split " + splitnumber.ToString() + " - " + type, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Sundays ISN Printer Split " + splitnumber.ToString() + " - " + type + ".png");

            }
            else
            {
                screencapture.printScreen("Sundays ISN Printer - " + type);
                mail.email_send("Sundays ISN Printer - " + type, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Sundays ISN Printer - " + type + ".png");
            }

            if (split)
            {

                if (splitnumber == 1)
                {
                    if (type == "Agencies")
                    {
                        runStatementReportBIWeekly("Weekly Bills", "A", "LedgerAgency_Weekly_" + DateTime.Today.AddMonths(-1).Year.ToString() + DateTime.Today.AddMonths(-1).Month.ToString().PadLeft(2, '0') + method.numberdays(DateTime.Today.AddMonths(-1)).ToString(), true);
                    }
                    if (type == "Ledger")
                    {
                        runStatementReportBIWeekly("Weekly Bills", "L", "Ledger_Weekly_"+DateTime.Today.AddMonths(-1).Year.ToString() + DateTime.Today.AddMonths(-1).Month.ToString().PadLeft(2,'0') + method.numberdays(DateTime.Today.AddMonths(-1)).ToString()  , true);
                    }
                }
                else
                {
                    if (type == "Agencies")
                    {
                        runStatementReportBIWeekly("Weekly Bills", "A", "LedgerAgency_Weekly_" + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0'), true);
                    }

                    if (type == "Ledger")
                    {
                        runStatementReportBIWeekly("Weekly Bills", "L", "Ledger_Weekly_" + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString().PadLeft(2, '0') + DateTime.Today.Day.ToString().PadLeft(2, '0'), true);
                    }
                }

            }
            
            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        }

    

       

       

        public void runEOMISN(string statkey, string client, string type)
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchISNPrinter();
            method.getButton(method.mainwind, "Generate Statements");
            Window invprinter = method.getWindow("Statement Printer");
            if (Convert.ToBoolean(ConfigurationManager.AppSettings["Maximize"]))
            {
                method.getButton(invprinter, "Maximize");
            }
            ComboBox company = method.getComboBoxByClass("TComboBox", invprinter, 0);
            company.Select("Oregonian");
            TextBox statementkey = method.getTextBox("", invprinter, 0);
            statementkey.BulkText = statkey;
            if (client == "O")
            {
                method.getCheckBox("Clear Table Data if Company Selected?", invprinter, 0);
            }
            else {
                method.getUnCheckBox("Clear Table Data if Company Selected?", invprinter, 0);
            }
            method.getButton(invprinter, "Limit to Customer Type");
            if (client == "O")
            {
                method.selectcustomerforrepISN("No");
                method.selectcustomerforrepISN("O");
                method.selectcustomerforrepISN("X");
                method.selectcustomerforrepISN("Y");
                method.selectcustomerforrepISN("Z");

            }
            else
            {
                method.selectcustomerforrepISN(client);
            }
            Window custwind = method.getWindow("Customer Type Selection");
            method.getButton(custwind, "OK");
            ComboBox disttype = method.getComboBoxByClass("TComboBox", invprinter, 2);
            disttype.Items.Clear();
            if(type=="Ledger" || type== "NonLedger")
            {
                method.getRadioButton("Exclude Agencies", invprinter, 0);
            }
            if (type == "Other") 
            {
                method.getRadioButton("All Customers", invprinter, 0);

            }
            if (type == "Agencies")
            {
                method.getRadioButton("Agencies Only Aging By Client", invprinter, 0);
            }
            method.getRadioButton("All Customers", invprinter, 1);
            ComboBox frequency = method.getComboBoxByClass("TComboBox", invprinter, 3);
            frequency.Items.Clear();
            ComboBox period = method.getComboBoxByClass("TComboBox", invprinter, 4);
            period.Select(PredefinedPeriodEOM());
            method.SetBalFirstDates(1, invprinter, -1);
            method.SetBalFirstDates(2, invprinter, 0);
            method.getUnCheckBox("Use Different Activity Date Range?", invprinter, 0);
            method.getUnCheckBox("Include Closed Transactions?", invprinter, 0);
            method.getCheckBox("Include Printed Invoices?", invprinter, 0);
            method.getRadioButton("Balance Forward", invprinter, 0);
            method.getCheckBox("Detail Line for Each Run Date?", invprinter, 0);
            method.getRadioButton("Actual Run Date Cost", invprinter, 0);
            method.getCheckBox("Separate Color Charges?", invprinter, 0);
            method.getCheckBox("Separate Position Premium?", invprinter, 0);
            method.getCheckBox("Separate Material Charges?", invprinter, 0);
            method.getCheckBox("Separate Discounts?", invprinter, 0);
            method.getCheckBox("Create Single Contract Discount Detail for Each Invoice?", invprinter, 0);
            method.getCheckBox("Details Typo Cost?", invprinter, 0);
            method.getCheckBox("Details Color Cost?", invprinter, 0);
            method.getCheckBox("Create Short Rate / Rebate Details?", invprinter, 0);
            method.getCheckBox("Separate Typography Details?", invprinter, 0);
            method.getRadioButton("Auto Generate-MM-DD-YYYY", invprinter, 0);
            method.getRadioButton("Customer Name", invprinter, 0);
            method.getCheckBox("Create Contract Data for Subrep?", invprinter, 0);
            method.getRadioButton("All through Period End - Only Open", invprinter, 0);
            method.SetBalEndDates(6, invprinter, -1);
            method.dotabs(1);
            KeyStrokes key = new KeyStrokes();
            key.sendvalue(PredefinedPeriodEOM());


            method.getButton(invprinter, "Create Data");
            method.app.WaitWhileBusy();
            try{
            if (invprinter.HasPopup())
            {
                Window warning = invprinter.ModalWindow("Warning");
                method.getButton(warning, "OK");
            }
                    }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);


                    }

            method.iddleStatePorgressBar(invprinter);

            method.app.WaitWhileBusy();
            screencapture.printScreen("EOM ISN Printer - "+type);

            mail.email_send("EOM ISN Printer - " + type, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "EOM ISN Printer - " + type + ".png");

            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }

        }


     

        public void ISNPrinterMIDM(string statkey, string client, string type)
        {
             string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchISNPrinter();
            method.getButton(method.mainwind, "Generate Statements");
            Window invprinter = method.getWindow("Statement Printer");
            if (Convert.ToBoolean(ConfigurationManager.AppSettings["Maximize"]))
            {
                method.getButton(invprinter, "Maximize");
            }
            ComboBox company = method.getComboBoxByClass("TComboBox", invprinter, 0);
            company.Select("Oregonian");
            TextBox statementkey = method.getTextBox("", invprinter, 0);
            statementkey.BulkText = statkey;
            if (client == "O")
            {
                method.getCheckBox("Clear Table Data if Company Selected?", invprinter, 0); 
            }
            else
            {
                method.getUnCheckBox("Clear Table Data if Company Selected?", invprinter, 0);
            }
            method.getButton(invprinter, "Limit to Customer Type");

            if (client == "O")
            {
                method.selectcustomerforrepISN("No");
                method.selectcustomerforrepISN("O");
                method.selectcustomerforrepISN("X");
                method.selectcustomerforrepISN("Y");
                method.selectcustomerforrepISN("Z");
            }
            else {

                method.selectcustomerforrepISN(client);
            
            }
            Window custwind = method.getWindow("Customer Type Selection");
            method.getButton(custwind, "OK");
            ComboBox disttype = method.getComboBoxByClass("TComboBox", invprinter, 2);
            disttype.Items.Clear();
            if (type == "Ledger" || type == "NonLedger")
            {
                method.getRadioButton("Exclude Agencies", invprinter, 0);
            }
            if (type == "Other")
            {
                method.getRadioButton("All Customers", invprinter, 0);

            }
            if (type == "Agencies")
            {
                method.getRadioButton("Agencies Only Aging By Client", invprinter, 0);
            }
            method.getRadioButton("All Customers", invprinter, 1);
            ComboBox frequency = method.getComboBoxByClass("TComboBox", invprinter, 3);
            frequency.Items.Clear();
            ComboBox period = method.getComboBoxByClass("TComboBox", invprinter, 4);
            period.Select(PredefinedPeriod());
            method.SetBalFirstDates(1, invprinter, 0);
            method.SetBalFirstDates(2, invprinter, 1);
            method.getCheckBox("Use Different Activity Date Range?", invprinter, 0);
            method.getCheckBox("Age using Differrent Acivity End Date?", invprinter, 0);
            method.SetBalFirstDates(3, invprinter, 0);
            method.SetDateTodayTomorrow(4, invprinter, 1);
            method.getUnCheckBox("Include Closed Transactions?", invprinter, 0);
            method.getCheckBox("Include Printed Invoices?", invprinter, 0);
            method.getRadioButton("Balance Forward", invprinter, 0);
            method.getUnCheckBox("Detail Line for Each Run Date?", invprinter, 0);
            method.getUnCheckBox("Separate Color Charges?", invprinter, 0);
            method.getUnCheckBox("Separate Position Premium?", invprinter, 0);
            method.getUnCheckBox("Separate Material Charges?", invprinter, 0);
            method.getUnCheckBox("Separate Discounts?", invprinter, 0);
            method.getUnCheckBox("Create Single Contract Discount Detail for Each Invoice?", invprinter, 0);
            method.getUnCheckBox("Details Typo Cost?", invprinter, 0);
            method.getUnCheckBox("Details Color Cost?", invprinter, 0);
            method.getUnCheckBox("Create Short Rate / Rebate Details?", invprinter, 0);
            method.getUnCheckBox("Separate Typography Details?", invprinter, 0);
            method.getRadioButton("No Statement Number", invprinter, 0);
            method.getRadioButton("Customer Name", invprinter, 0);
            method.getUnCheckBox("Create Contract Data for Subrep?", invprinter, 0);

            method.SetDateTodayTomorrow(6, invprinter, 0);
            method.dotabs(1);
            KeyStrokes key = new KeyStrokes();
            key.sendvalue(PredefinedPeriodMidM());

            method.getButton(invprinter, "Create Data");
            method.app.WaitWhileBusy();
          try{
            if (invprinter.HasPopup())
            {
                Window warning = invprinter.ModalWindow("Warning");
                method.getButton(warning, "OK");
            }
                    
              }
                    catch (Exception ex)
                    {

                        log = ex.Message.ToString();
                        logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                      

                    }

            method.iddleStatePorgressBar(invprinter);

            method.app.WaitWhileBusy();
            screencapture.printScreen("Mid-Month ISN Printer - "+type);

            mail.email_send("Mid-Month ISN Printer - " + type, ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Mid-Month ISN Printer - " + type + ".png");

            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }

        }


        public void runBalanceUtility() {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchBalanceUtility();

            method.SetDateTodayTomorrow(0, method.mainwind, 0);
            method.SetDateTodayTomorrow(1, method.mainwind, 1);
            method.getCheckBox("Ignore Range - Check All Transactions", method.mainwind, 0);           
            method.SetDateTodayTomorrow(2, method.mainwind, 0);
            method.getCheckBox("Check Payments", method.mainwind, 0);
            method.getCheckBox("Check Invoices", method.mainwind, 0);
            method.getCheckBox("Check Credits", method.mainwind, 0);
            method.getCheckBox("Check Debits", method.mainwind, 0);
            method.getCheckBox("Check Relationships?", method.mainwind, 0);
            method.getUnCheckBox("Include Closed Transactions?", method.mainwind, 0);
            method.getCheckBox("Execute SQL Transactions", method.mainwind, 0);
            method.SetDateTodayTomorrow(4, method.mainwind, 0);
            method.SetDateTodayTomorrow(3, method.mainwind, 0);
            method.SetDateTodayTomorrow(5, method.mainwind, 1);

            method.getButton(method.mainwind, "Check");

            method.isProcessIddle("OracleBalanceUtility");

            screencapture.printScreen("Balance Utility 1");

            mail.email_send("Balance Utility - Part 1", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Balance Utility 1.png");

            method.getCheckBox("Close Transactions?", method.mainwind, 0);

            method.getButton(method.mainwind, "Check");

            method.isProcessIddle("OracleBalanceUtility");

            screencapture.printScreen("Balance Utility 2");

            mail.email_send("Balance Utility - Part 2", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Balance Utility 2.png");

            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }

        }
       

        private string PredefinedPeriodMidM() {
            string period;

            period = DateTime.Today.Month + "/1/" + DateTime.Today.Year + " - " + DateTime.Today.Month + "/15/" + DateTime.Today.Year;
            return period;
        }

        private string PredefinedPeriodEOM() {
            string period;
            if (DateTime.Today.Month == 1)
            {
                period = DateTime.Today.AddMonths(-1).Month + "/1/" + DateTime.Today.AddYears(-1).Year + " - " + DateTime.Today.AddMonths(-1).Month + "/" + DateTime.DaysInMonth(DateTime.Today.AddYears(-1).Year, DateTime.Today.AddMonths(-1).Month) + "/" + DateTime.Today.AddYears(-1).Year;

            }
            else
            {
                period = DateTime.Today.AddMonths(-1).Month + "/1/" + DateTime.Today.Year + " - " + DateTime.Today.AddMonths(-1).Month + "/" + DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.AddMonths(-1).Month) + "/" + DateTime.Today.Year;
            }
            return period;
        }

        private string PredefinedPeriod() {

            string period;

            period = DateTime.Today.Month + "/1/" + DateTime.Today.Year + " - " + DateTime.Today.Month + "/" + DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month) + "/" + DateTime.Today.Year;

           
            return period;
        
        }

        private string PredefinedPeriodSplit()
        {

            string period;
            string year = "";

            if (DateTime.Today.AddMonths(-1).Month.ToString() == "12")
            {

                year = DateTime.Today.AddYears(-1).Year.ToString();
            }
            else {

                year = DateTime.Today.Year.ToString();
            
            }

            period = DateTime.Today.AddMonths(-1).Month + "/1/" + year + " - " + DateTime.Today.AddMonths(-1).Month + "/" + DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.AddMonths(-1).Month) + "/" + year;


            return period;

        }

        public void runStatementReportBIWeekly(string distcode, string letter, string reportname, bool multivalue) 
        {
             string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            CrystalReports statementrep = new CrystalReports();
            statementrep.CR(distcode,letter, reportname, multivalue);
            mail.email_send("Statement Report", "");

            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
         
        }
        public void visitLockAdmin()
        {
            string log = "";
            int loop = 0;
            TabPage page = method.getTabPage("Locks");
            TextBox text = method.getTextBoxByClass("", method.mainwind, 0);
          
            loop = Convert.ToInt32(text.Name.Substring(0, Convert.ToInt32(text.Name.IndexOf(" "))));
           
            method.moveclickmouse(page, 19, 77);

            method.LockAdminLoops(loop);
           
            method.moveclickmouse(page, 40, -25);
            try{
            if (method.mainwind.HasPopup())
            {
                Window warning = method.mainwind.ModalWindow("Warning");
                method.getButton(warning, "OK");
            }

            }
            catch (Exception ex)
            {

                log = ex.Message.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);


            }

        }



        public void runAgingUtility() 
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchAgingUtility();
            method.dotabs(5);
            KeyStrokes key = new KeyStrokes();
            key.sendvalue("{DOWN}");
            key.sendvalue("{DOWN}");
            method.dotabs(2);
            key.sendspace();
            method.dotabs(1);
            key.sendvalue("{DOWN}");
            method.dotabs(1);
            key.sendspace();
            int times = 2;
            while (times > 0)
            {
                method.dotabs(1);
                key.sendvalue(DateTime.Today.Month.ToString());
                key.sendvalue("{RIGHT}");
                key.sendvalue(DateTime.Today.Day.ToString());
                key.sendvalue("{RIGHT}");
                key.sendvalue(DateTime.Today.Year.ToString());
                times -= 1;
            }
            method.dotabs(2);
            key.sendspace();
            method.dotabs(1);
            key.sendspace();
            method.dotabs(5);


            key.sendvalue("{ENTER}");
            try
            {
                if (method.mainwind.HasPopup())
                {
                    Window warning = method.mainwind.ModalWindow("Warning");
                    method.getButton(warning, "OK");
                }
            }
            catch (Exception ex)
            {

                log = ex.Message.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);


            }
          

            method.isProcessIddle("AgingUtility");

            method.iddleState("AgingUtility");

            method.app.WaitWhileBusy();

            screencapture.printScreen("Aging Utility - Weekly");

            mail.email_send("Aging Utility - Weekly", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Aging Utility - Weekly" + ".png");

            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        }


        public void runFinanceManager() 
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchFinanceManager();
            MenuBar menu=method.getMenuBar(method.mainwind, "Application");
            method.clickMenuItem(menu, "Options");
            method.clickMenuItem(menu, "Fiscal Year...");
            method.app.WaitWhileBusy();
            Window fiscalyears = method.getWindow("Fiscal Years");
            ListView table = method.getDataGrid(fiscalyears);
            AutomationElement element = null;
            if (DateTime.Today.Month == 1)
            {
                element = table.GetElement(SearchCriteria.ByText(DateTime.Today.AddYears(-1).Year.ToString()));
            }
            else {

                element = table.GetElement(SearchCriteria.ByText(DateTime.Today.Year.ToString()));
            }
            element.SetFocus();
            method.moveclickmouse(element, 0, 0);
            method.getButton(fiscalyears, "Edit");
            Window fiscalyear = method.getWindow("Fiscal Year");
            ListView table2 = method.getDataGrid(fiscalyear);
            int month = DateTime.Now.AddMonths(-1).Month;
            if(month>9){

                for (int i = 0; i < 3; i++)
                {
                    table2.ScrollBars.Vertical.ScrollDown();
                }
            }

           
              AutomationElement element2 = table2.GetElement(SearchCriteria.ByText(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.AddMonths(-1).Month)));
          
             
            element2.SetFocus();
            method.moveclickmouse(element2, 0, 0);
            method.getButton(fiscalyear, "Edit");
            method.app.WaitWhileBusy();
            Window acctperiod = method.getWindow("Accounting Period");
            ComboBox status=method.getComboBoxByClass("TComboBox", acctperiod, 0);
            status.Select("Closed");
            method.app.WaitWhileBusy();
            KeyStrokes key = new KeyStrokes();
            key.sendvalue("{ENTER}");

            try
            {
                if (acctperiod.HasPopup())
                {
                    Window warning = acctperiod.ModalWindow("Warning");
                    method.getButton(warning, "OK");
                    method.app.WaitWhileBusy();
                }
                if (acctperiod.HasPopup())
                {
                    Window warning = acctperiod.ModalWindow("Error");
                    method.getButton(warning, "OK");
                    method.app.WaitWhileBusy();

                    Window unrolledadj = method.getWindow("Unrolled Adjustments");
                    method.getButton(unrolledadj, "OK");
                    method.app.WaitWhileBusy();
                }
                if (acctperiod.HasPopup())
                {
                    Window warning = acctperiod.ModalWindow("Warning");
                    method.getButton(warning, "OK");
                    method.app.WaitWhileBusy();
                }
            }
            catch (Exception ex)
            {

                log = ex.Message.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);

            }

              try
              {
            if (acctperiod.Visible)
            {
                method.getButton(acctperiod, "Apply");
            }
           
            method.app.WaitWhileBusy();
              }
                     catch (Exception ex)
                     {

                         log = ex.Message.ToString();
                         logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);

                     }

              try
                {
            if (acctperiod.HasPopup())
            {
              
                    Window warning = acctperiod.ModalWindow("Warning");
                    method.getButton(warning, "OK");
               

            }
                }
              catch (Exception ex)
              {

                  log = ex.Message.ToString();
                  logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);

              }

              try
              {
                  if (acctperiod.Visible)
                  {
                      method.getButton(acctperiod, "Apply");
                  }
              }
              catch (Exception ex)
              {

                  log = ex.Message.ToString();
                  logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);

              }
                     try
              {
                  method.app.WaitWhileBusy();
                  method.getButton(fiscalyear, "Apply");
                  method.app.WaitWhileBusy();
                  method.getButton(fiscalyears, "Done");
                  method.app.WaitWhileBusy();
                  method.clickMenuItem(menu, "File");
                  method.app.WaitWhileBusy();
                  method.clickMenuItem(menu, "Exit");
                  method.app.WaitWhileBusy();
              }
                     catch (Exception ex)
                     {

                         log = ex.Message.ToString();
                         logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);

                     }

            mail.email_send("Finance Manager", "");

            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        }


       

        public void runOracleReOpenTool() {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]))
            {
                try
                {
            method.launchOracleReOpenTool();
            method.getButton(method.mainwind, "Run");
            method.iddleStatePorgressBar(method.mainwind);

            method.isProcessIddle("OracleReopenTransactionsTool");

            screencapture.printScreen("Oracle Reopen Tool");

            mail.email_send("Oracle Reopen Tool", ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + "Oracle Reopen Tool" + ".png");

            KillRunningProcesses();
            tries = Convert.ToInt32(ConfigurationManager.AppSettings["StepTries"]);
                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
                    KillRunningProcesses();
                    tries += 1;

                }
            }
        
        }

        public void KillRunningProcesses() {
            
            method.killProcess("AcroRd32");
            method.killProcess("Acrobat");
            method.killProcess("OracleReports");
            method.killProcess("LockAdmin");
            method.killProcess("InvoiceGenerator");
            method.killProcess("OracleISNPrinter");
            method.killProcess("notepad");
            method.killProcess("MemoBill-Send");
            method.killProcess("AgingUtility");
            method.killProcess("FinanceManager");
            method.killProcess("OracleReopenTransactionsTool");
            method.killProcess("OracleBalanceUtility");
            method.killProcess("PrimoPDF");
        }
     

    }
}
