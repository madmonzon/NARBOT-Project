using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TestStack.White;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.MenuItems;
using TestStack.White.UIItems;
using TestStack.White.Factory;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.ListBoxItems;
using TestStack.White.InputDevices;
using System.Windows;
using TestStack.White.UIItems.PropertyGridItems;
using System.Diagnostics;
using TestStack.White.AutomationElementSearch;
using System.Globalization;
using System.Configuration;
using TestStack.White.UIItems.TreeItems;
using TestStack.White.UIItems.TabItems;
using TestStack.White.UIItems.WindowStripControls;
using TestStack.White.UIItems.TableItems;
using System.Windows.Automation;
using System.Windows.Interop;

namespace NARBOT
{
    class Core
    {
        [System.Runtime.InteropServices.DllImport("user32.d ll")]
        public static extern int SetWindowPos(IntPtr hwnd, IntPtr
        hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);

        [System.Runtime.InteropServices.DllImport("user32.d ll")]
        public static extern int BringWindowToTop(IntPtr hwnd);

        public IntPtr HWND_TOPMOST = (IntPtr)(-1);
        public IntPtr HWND_NOTOPMOST = (IntPtr)(-2);
        public int SWP_NOSIZE = 0x1;

        WriteLog logfile = new WriteLog();
        public Application app;
        public Window mainwind;
        private ScreenShots screencapture = new ScreenShots();
        //public EMail mail = new EMail();
       
        //*********************** METHODS ******************************

        public void launchRemoteCommand()
        {
            app = Application.Launch(ConfigurationManager.AppSettings[@"RemoteCommand"]);


            mainwind = getWindow("Desktop Central Free Windows Tools");
        }



        //public void launchMemoBills() 
        //{
        //    Process proc;
        //        //Process.Start(ConfigurationManager.AppSettings[@"MemoBillsEXE"]);

        //    proc = new Process
        //    {
        //        StartInfo = new ProcessStartInfo
        //        {
        //            FileName = ConfigurationManager.AppSettings[@"MemoBillsBat"],
        //            UseShellExecute = false,
        //            RedirectStandardOutput = true,
        //            CreateNoWindow = true,

        //        }

        //    };
        
        //    proc.Start();
        //    proc.StandardOutput.ReadToEnd();
        //}

        public void launchAgingUtility()
        {
            app = Application.Launch(ConfigurationManager.AppSettings[@"AgingUtilPath"]);
            Loginapp(false);

            mainwind = getWindow("Aging Utility");
        }


        public void launchInvoiceGenerator()
        {
            app = Application.Launch(ConfigurationManager.AppSettings[@"InvoiceGenerator"]);
            Loginapp(true);

            mainwind = getWindow("Invoice Generator   User=igstore   DB=" + ConfigurationManager.AppSettings["DefaultDatabase"]);
           
        }

        public void launchISNPrinter()
        {
            app = Application.Launch(ConfigurationManager.AppSettings[@"ISNPrinter"]);
            Loginapp(false);

            mainwind = getWindow("AdBase Invoice/Statement/Notice Printer");
            
        }


        public void launchlockadmin() {
            app = Application.Launch(ConfigurationManager.AppSettings[@"LockAdmin"]);
            Loginapp(false);
       
            mainwind = getWindow("Lock Admin");
        }

        public void launchOracleReports() { 
        
         app = Application.Launch(ConfigurationManager.AppSettings[@"OracleReportsDIR"]);
         Loginapp(false);
       
         mainwind = getWindow("Standard Reports");
        }



        public void launchOracleReOpenTool()
        {

            app = Application.Launch(ConfigurationManager.AppSettings[@"OracleReOpenTool"]);
         
            Loginapp(false);

            mainwind = getWindow("Oracle Reopen Transactions Tool");
        }

        public void launchBalanceUtility()
        {
           
            //Process.Start(ConfigurationManager.AppSettings[@"BalanceUtility"]);
            //System.Threading.Thread.Sleep(60000);
            //mainwind = getWindow("C:\\Windows\\system32\\cmd.exe");

            app = Application.Launch(ConfigurationManager.AppSettings[@"BalanceUtility"]);

            Loginapp(false);

            mainwind = getWindow("Balance Utility");
        }

        public void launchFinanceManager()
        {

            app = Application.Launch(ConfigurationManager.AppSettings[@"FinanceManagerPath"]);
            Loginapp(false);

            mainwind = getWindow("Finance Manager");
        }

        //Login Application
        public void Loginapp(bool igstore)
        {
 
            Window wind = getWindow("Login");


            var textboxu = getTextBox("", wind, 0);
            var textboxp = getTextBox("", wind, 1);
            if (igstore) {
                textboxu.BulkText = ConfigurationManager.AppSettings["UsernameIGStore"];
                textboxp.BulkText = ConfigurationManager.AppSettings["PasswordIGStore"];
            }
            else
            {
                textboxu.BulkText = ConfigurationManager.AppSettings["Username"];
                textboxp.BulkText = ConfigurationManager.AppSettings["Password"];
            }
            getButton(wind, "OK");
            app.WaitWhileBusy();
           

        }

        //Closes any window
        public void closewindow(Window wind) {

            wind.Close();
        
        }

       
        //printout the report as a PDF file
        public void printpreviewwindow(string path, string fname)
        {
            try
            {
                Window printpreview = getWindow("Print Preview");
                app.WaitWhileBusy();
                getButtonByID(printpreview, "Item 10");

                Window dialogprint = getWindow("Print");
                ComboBox cbox = getComboBoxByClass("ComboBox", dialogprint, 0);
                cbox.Select(ConfigurationManager.AppSettings["DefaultPrinter"]);

                getButton(dialogprint, "OK");

                getButtonByID(printpreview, "Item 11");

                app.WaitWhileBusy();

                check();

                if (ConfigurationManager.AppSettings["DefaultPrinter"].ToString() == "PrimoPDF")
                {
              
                        runprimoPDF(path, fname);
              
                }



                closewindow(printpreview);
            }
            catch (Exception ex)
            {

                WriteLog logger = new WriteLog();
                logger.writefile(ex.Message.ToString(), ConfigurationManager.AppSettings["LogFile"]);
                screencapture.printScreen(fname);
            }
            
        }


        public void runprimoPDF(string path, string fname)
        {

            app.WaitWhileBusy();
            createPath(path);
            attachprimopdf();
            Window primopdfwindow = getWindow("PrimoPDF by Nitro PDF Software");
           
            getButton(primopdfwindow, "Create PDF");
          
            Window saveas = primopdfwindow.ModalWindow("Save As");
            
            TextBox savefile = getTextBoxByClass("Edit", saveas, 0);
      
            KeyStrokes key = new KeyStrokes();

            if (Convert.ToBoolean(ConfigurationManager.AppSettings["SaveAs_WindowNV"]))
            {
                
                dotabs(5);
                
                key.sendvalue("{ENTER}");
               
                key.sendvalue(path);
               
                key.sendvalue("{ENTER}");
              
                savefile.Focus();
                key.sendvalue(fname);
            }

            if (Convert.ToBoolean(ConfigurationManager.AppSettings["SaveAs_WindowOV"]))
            {

              
                savefile.Focus();
                key.sendvalue(path + fname + ".pdf");

            }
           
            getButton(saveas, "Save");

            try
            {
                app.WaitWhileBusy();

                iddleState("PrimoPDF");

                bool found = false;
                int count = 0;

                while (!found)
                {
                    
                    List<Window> modalWindows = primopdfwindow.ModalWindows();
                    count += 1;
                    System.Threading.Thread.Sleep(1000);
                    foreach (Window mywindow in modalWindows)
                    {
                        
                        if (mywindow.Name == "PrimoPDF by Nitro PDF")
                        {
                            
                            Window overwrite = primopdfwindow.ModalWindow("PrimoPDF by Nitro PDF");
                          
                            getButton(overwrite, "Overwrite");
                            app.WaitWhileBusy();
                            found = true;

                        }
                    }

                    if (count == Convert.ToInt32(ConfigurationManager.AppSettings["WaitForOW"]))
                    {

                        found = true;

                    }
                }
            }
            catch (Exception ex)
            {

                WriteLog logger = new WriteLog();
                logger.writefile(ex.Message.ToString(), ConfigurationManager.AppSettings["LogFile"]);

            }

        }

      

        //Attach PrimoPDF process to the current runing application 
        public void attachprimopdf()
        {
            bool notfound = true;
            while (notfound)
            {
                Process[] pname = Process.GetProcessesByName("PrimoPDF");
                foreach (Process myprocess in pname)
                {
                    app = Application.Attach(myprocess.Id);
                    notfound = false;
                    
                }

            }
            app.WaitWhileBusy();
           
        }
       
        public void attachMemoBills()
        {
            bool notfound = true;
            while (notfound)
            {
                Process[] pname = Process.GetProcessesByName("MemoBill-Send");
                foreach (Process myprocess in pname)
                {
                    app = Application.Attach(myprocess.Id);
                    notfound = false;
                }

            }
            app.WaitWhileBusy();
        }

        public TitleBar getTitleBar(Window wind, string titlebar) 
        {

            TitleBar title = wind.Get<TitleBar>(SearchCriteria.ByText(titlebar));

            return title;
        
        }

        public void getNPADMenuItem(Window wind) {

          MenuBar bar= wind.GetMenuBar(SearchCriteria.ByText("Application"));
       
        var item=   bar.MenuItem("File");
        item.Click();
        KeyStrokes k = new KeyStrokes();
        for (int i = 0; i < 4; i++)
        {
            k.sendvalue("{DOWN}");
        }
        k.sendvalue("{ENTER}");
          
        }

        public void saveLogFile(string path, Window np) {

            Window saveas =np.ModalWindow("Save As");
           
            TextBox savefile = getTextBoxByClass("Edit", saveas, 0);
            string fname = savefile.Text;
            KeyStrokes key = new KeyStrokes();
          

            if (Convert.ToBoolean(ConfigurationManager.AppSettings["SaveAs_WindowNV"]))
            {

               dotabs(6);

               key.sendvalue("{ENTER}");

                key.sendvalue(ConfigurationManager.AppSettings[@"InvoiceGenLog"]);
               key.sendvalue("{ENTER}");

               savefile.Focus();
               key.sendvalue(fname);
            }

            if (Convert.ToBoolean(ConfigurationManager.AppSettings["SaveAs_WindowOV"]))
            {


                savefile.Focus();
                key.sendvalue(path + @"\" + fname + ".txt");

            }
            key.sendvalue("{ENTER}");


           if(saveas.HasPopup()){
              Window confirm=saveas.ModalWindow("Confirm Save As");
              getButton(confirm, "Yes");
            }
        }

        public void attachnotepad()
        {
          
            bool notfound = true;
            while (notfound)
            {
                Process[] pname = Process.GetProcessesByName("notepad");
                foreach (Process myprocess in pname)
                {
                   
                    app = Application.Attach(myprocess.Id);
                    notfound = false;
                    try
                    {
                     
                        Window npadinfo = getWindow("InvoiceGeneratorInfoLog - Notepad");
                        npadinfo.Focus();
                        getNPADMenuItem(npadinfo);
                        saveLogFile("", npadinfo);
                    }
                    catch { }
                }

            }
            app.WaitWhileBusy();

          
        }

        public void iddleState(string name) {
            Process[] pname = Process.GetProcessesByName(name);
            foreach (Process myprocess in pname)
            {
                
                myprocess.WaitForInputIdle();

               

            }
        }

       public bool check12AM()
        {

            bool check = false;

            if (DateTime.Now.ToLongTimeString().Contains("AM"))
            {

                check = true;

            }


            return check;
        }

        public void iddleStatePorgressBar(Window wind)
        {
            ProgressBar b = FindProgressBar(wind);
            int count = 0;
            string previousvalue = "";
            while (count < Convert.ToInt32(ConfigurationManager.AppSettings["ProgressBarWaitTime"]))
            {

                if (b.Value.ToString() != previousvalue)
                {
                    previousvalue = b.Value.ToString();
                    count = 0;
                   app.WaitWhileBusy();
                    System.Threading.Thread.Sleep(1000);


                }
                else
                {

                    count += 1;
                   app.WaitWhileBusy();
                    System.Threading.Thread.Sleep(1000);
                }
            }
        }

        public MenuBar getMenuBar(Window wind, string name) 
        { 
        
       MenuBar menu= wind.Get<MenuBar>(SearchCriteria.ByText(name));

       return menu;
        
        }
       
        private void check() {

            try
            {
              
                    Window wind = app.GetWindow("Printing progress");
                    while (wind.Visible)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                
            }
            catch (Exception ex) {

                WriteLog logger = new WriteLog();
                logger.writefile(ex.Message.ToString(), ConfigurationManager.AppSettings["LogFile"]);
            
            }
        
        }

        public ProgressBar FindProgressBar(Window wind)
        {

            ProgressBar bar = wind.Get<ProgressBar>(SearchCriteria.ByClassName("TProgressBar"));
            app.WaitWhileBusy();
            return bar;
        }

        //Move the cursor by doing tabs on a window. Number of tabs is expected
        public void dotabs(int tabnum) {
            KeyStrokes tab = new KeyStrokes();
            while (tabnum > 0)
            {
                tab.sendvalue("{TAB}");
                tabnum -= 1;
            }
        
        }

       

        //Kills any process if is running 
        public void killProcess(string process)
        {
            try
            {
                Process[] pname = Process.GetProcessesByName(process);
                foreach (Process myprocess in pname)
                {
                   
                    myprocess.Kill();
                }
            }
            catch (Exception ex) {

                WriteLog logger = new WriteLog();
                logger.writefile(ex.Message.ToString(), ConfigurationManager.AppSettings["LogFile"]);
            }
        }

        public void SetDayOfWeek(int i, Window wind, DateTime date) 
        {
            var prevbaldate = getPanel(i, wind);
            prevbaldate.Focus();
            datetimewrite(Convert.ToString(date.Month), Convert.ToString(date.Day), Convert.ToString(date.Year));
        
        }
        //Sets the dates for end of the month in any date time picker
        public void SetBalEndDates(int i, Window wind, int months)
        {
            int nd;
            DateTime date;
            var prevbaldate = getPanel(i, wind);
            prevbaldate.Focus();
            date = DateTime.Today;
            date = date.AddMonths(months);
            nd = numberdays(date);
           
                datetimewrite(Convert.ToString(date.Month), Convert.ToString(nd), Convert.ToString(date.Year));
         
        }
        //Sets the dates for first of the month in any date time picker
        public void SetBalFirstDates(int i, Window wind, int months)
        {
            int nd;
            DateTime date;
            var prevbaldate = getPanel(i, wind);
            prevbaldate.Focus();
            date = DateTime.Today;
            date = date.AddMonths(months);
            nd = 1;
          
                datetimewrite(Convert.ToString(date.Month), Convert.ToString(nd), Convert.ToString(date.Year));
          
        }


        public void SetBalFirstDatesSplit(int i, Window wind, int months)
        {
            int nd;
            DateTime date;
            var prevbaldate = getPanel(i, wind);
            prevbaldate.Focus();
            date = DateTime.Today;
            date = date.AddMonths(months);
            nd = 1;
           
                datetimewrite(Convert.ToString(date.Month), Convert.ToString(nd), Convert.ToString(date.Year));
           
           
        }

        public void SetBalFirstDatesCHK(int i, Window wind, int months)
        {
            int nd;
            DateTime date;
            var prevbaldate = getPanel(i, wind);
            prevbaldate.Focus();
            date = DateTime.Today;
            date = date.AddMonths(months);
            nd = 1;
           
                datetimewritechk(Convert.ToString(date.Month), Convert.ToString(nd), Convert.ToString(date.Year));
          
            
        }
       

        //Sets the dates for today and tomorrow
        public void SetDateTodayTomorrow(int i, Window wind, int days)
        {
            
            DateTime date;
            var prevbaldate = getPanel(i, wind);
            prevbaldate.Focus();
            date = DateTime.Today;
            int day = date.Day+days;
            if (days != -1)
            {
                if (days != 0 && date.Day == DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month))
                {

                    day = 1;

                    if (DateTime.Today.Month == 12)
                    {
                        datetimewrite(Convert.ToString(date.AddMonths(1).Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.AddYears(1).Year));
                    }
                    else
                    {

                        datetimewrite(Convert.ToString(date.AddMonths(1).Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.Year));
                    }

                }
                else
                {
                    datetimewrite(Convert.ToString(date.Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.Year));
                }
            }
            else
            {
                if (date.Day == 1)
                {

                    datetimewrite(Convert.ToString(date.Month + days).PadLeft(2, '0'), Convert.ToString(DateTime.DaysInMonth(date.AddDays(days).Year, date.AddDays(days).Month)).PadLeft(2, '0'), Convert.ToString(date.Year));
                }
                else 
                {
                    datetimewrite(Convert.ToString(date.Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.Year));
                }
            }
            
        }

        public void SetDateTodayTomorrowCHK(int i, Window wind, int days)
        {

            DateTime date;
            var prevbaldate = getPanel(i, wind);
            prevbaldate.Focus();
            date = DateTime.Today;

            int day = date.Day + days;
            if (days != -1)
            {
                if (days != 0 && date.Day == DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month))
                {

                    day = 1;

                    if (DateTime.Today.Month == 12)
                    {
                        datetimewritechk(Convert.ToString(date.AddMonths(1).Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.AddYears(1).Year));
                    }
                    else
                    {

                        datetimewritechk(Convert.ToString(date.AddMonths(1).Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.Year));
                    }

                }
                else
                {
                    datetimewritechk(Convert.ToString(date.Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.Year));
                }
            }
            else
            {
                datetimewrite(Convert.ToString(date.Month).PadLeft(2, '0'), Convert.ToString(day).PadLeft(2, '0'), Convert.ToString(date.Year));
            }
        }

        //Gets datetimepickers
        public Panel getPanel(int i, Window wind) {

        Panel pane = wind.Get<Panel>(SearchCriteria.ByClassName("TDateTimePicker").AndIndex(i));
        app.WaitWhileBusy();
        return pane;

        }

        //Finds checkboxes in a given window
        public void FindCheckbox(Window wind, string checkboxname) {

            CheckBox checkb = wind.Get<CheckBox>(SearchCriteria.ByText(checkboxname));
            app.WaitWhileBusy();
            checkb.Checked = true; 
        }
        //*********************buttons************************
       
        //Finds a button by name in a given window
        public void getButton(Window wind, string btnname)
        {
            Button btn = wind.Get<Button>(SearchCriteria.ByText(btnname));
            app.WaitWhileBusy();
            btn.Click();
        }

        //Finds a button by automation ID in a given window
        public void getButtonByID(Window wind, string btnautomationid)
        {
                bool notenable=true;
                Button btn = wind.Get<Button>(SearchCriteria.ByAutomationId(btnautomationid));
               
            while(notenable){
                app.WaitWhileBusy();
                if (btn.Enabled)
                {
                    btn.Click();
                    notenable = false;
                    
                }
               
                 }
               
        }
        
        public void gerptname(Window wind, int pos) {

            ListBox rptname = wind.Get<ListBox>(SearchCriteria.ByText("").AndIndex(0));
            app.WaitWhileBusy();
            rptname.Select(pos);
        }

      

        public Window getWindow(string title) {
           
      
            Window wind = app.GetWindow(title, InitializeOption.NoCache);
            app.WaitWhileBusy();
            return wind;
        
        }

        public int numberdays(DateTime date) {

            int numberday = DateTime.DaysInMonth(date.Year, date.Month);

            return numberday;
        }

        public void datetimewrite(string month, string day, string year) {

            KeyStrokes k = new KeyStrokes();
            k.sendvalue(month);
            k.sendvalue("{RIGHT}");
            k.sendvalue(day);
            k.sendvalue("{RIGHT}");
            k.sendvalue(year);
        }


        public void datetimewritechk(string month, string day, string year)
        {

            KeyStrokes k = new KeyStrokes();
            k.sendvalue("{RIGHT}");
            k.sendvalue(month);
            k.sendvalue("{RIGHT}");
            k.sendvalue("{RIGHT}");
            k.sendvalue(day);
            k.sendvalue("{RIGHT}");
            k.sendvalue("{RIGHT}");
            k.sendvalue("{RIGHT}");
            k.sendvalue(year);
        }


        public void eraseCustomerType(Window report, string name, int index, int erasex, int selectx)
        {
            var ctype = getTextBox(name, report, index);

            moveclickmouse(ctype, erasex, 0);
            moveclickmouse(ctype, selectx, 0);
        }

        //Gets the first letters in any listbox item
        public void firsLettersListbox(int i, ListBox listboxcust, KeyStrokes k, string lastones, int lenght)
        {
            if (lastones.Contains(listboxcust.Item(i).Name.Substring(0, lenght)))
            {
                
                k.sendspace();

            }
        
        }
            
        //Gets the customers that are used to run the report
        public void selectcustomerforrep(string custtype)
        {
            
            Window wind = app.GetWindow("Select Customer Types For Report", InitializeOption.NoCache);
            var listboxcust = getListBox("", wind, 0);  
            for (int i = listboxcust.Items.Count - 1; i >= 0; i--)
            {

                listboxcust.Check(listboxcust.Item(i).Name);
                KeyStrokes k = new KeyStrokes();
                
                if(custtype=="All" ){

                    k.sendspace();
                   
                   
                }
                else if (custtype == "L" )
                {
                    firsLettersListbox(i, listboxcust, k, custtype,1);
                  
                  

                }
                else if (custtype == "NL") {

                    firsLettersListbox(i, listboxcust, k, custtype,2);
                   
                
                }
                else if (custtype == "LA-LL-LM")
                {
                   
                    firsLettersListbox(i, listboxcust, k, custtype, 2);
                 

                }
                else if (custtype == "LG")
                {

                    firsLettersListbox(i, listboxcust, k, custtype, 2);


                }
                else if (custtype == "O")
                {

                    firsLettersListbox(i, listboxcust, k, custtype, 1);


                }
                else if (custtype == "X")
                {

                    firsLettersListbox(i, listboxcust, k, custtype, 1);


                }
                else if (custtype == "Y")
                {

                    firsLettersListbox(i, listboxcust, k, custtype, 1);


                }
                else if (custtype == "Z")
                {

                    firsLettersListbox(i, listboxcust, k, custtype, 1);


                }
                else if (custtype == "No-O-X-Y-Z-")
                {

                    firsLettersListbox(i, listboxcust, k, custtype, 2);


                }
                }

            }

        public void selectcustomerforrepISN(string custtype)
        {

            Window wind = app.GetWindow("Customer Type Selection", InitializeOption.NoCache);
            var listboxcust = getListBox("", wind, 0);
            for (int i = listboxcust.Items.Count - 1; i >= 0; i--)
            {

                listboxcust.Check(listboxcust.Item(i).Name);
                KeyStrokes k = new KeyStrokes();

                if (custtype == "L")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 1);



                }
                else if (custtype == "NL-B,NL-G")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 4);



                }
                else if (custtype == "NL")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 2);



                }
                else if (custtype == "No")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 2);



                }
                else if (custtype == "O")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 1);



                }
                else if (custtype == "X")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 1);



                }
                else if (custtype == "Y")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 1);



                }
                else if (custtype == "Z")
                {
                    firsLettersListbox(i, listboxcust, k, custtype, 1);



                }
            
            
            }

        }

        //Gets any textbox by name in a given window
        public TextBox getTextBox(string name, Window wind, int index) {

            TextBox textbox = wind.Get<TextBox>(SearchCriteria.ByText(name).AndIndex(index));
            app.WaitWhileBusy();
            return textbox;

        }

        //Gets any textbox by Class in a given window
        public TextBox getTextBoxByClass(string classname, Window wind, int index)
        {

            TextBox textbox = wind.Get<TextBox>(SearchCriteria.ByClassName(classname).AndIndex(index));
            app.WaitWhileBusy();
            return textbox;

        }

        //Gets any combobox by class in a given window
        public ComboBox getComboBoxByClass(string classname, Window wind, int index)
        {

            ComboBox combobox = wind.Get<ComboBox>(SearchCriteria.ByClassName(classname).AndIndex(index));
            app.WaitWhileBusy();
            return combobox;

        }

        //Gets any Listbox by name in a given window
        public ListBox getListBox(string name, Window wind, int index)
        {

            ListBox lbox = wind.Get<ListBox>(SearchCriteria.ByText(name).AndIndex(index));
            app.WaitWhileBusy();
            return lbox;

        }

        public void getRadioButton(string name, Window wind, int index)
        {

            RadioButton radio = wind.Get<RadioButton>(SearchCriteria.ByText(name).AndIndex(index));
            app.WaitWhileBusy();      
                radio.Select();

        }

        public void getCheckBox(string name, Window wind, int index)
        {

            CheckBox check = wind.Get<CheckBox>(SearchCriteria.ByText(name).AndIndex(index));
            app.WaitWhileBusy();
            check.Select();

        }
        public void getUnCheckBox(string name, Window wind, int index)
        {

            CheckBox check = wind.Get<CheckBox>(SearchCriteria.ByText(name).AndIndex(index));
            app.WaitWhileBusy();
            check.UnSelect();

        }
        //Moves the mouse to a location after finding the preceding control
        public void moveclickmouse(UIItem control, double Xplus, double Yplus) {

            double x = control.Location.X + Xplus;
            double y = control.Location.Y+Yplus;
            MouseClick m = new MouseClick();
            m.DoMouseClick((uint)x, (uint)y);
        }
        public void moveclickmouse(AutomationElement control, double Xplus, double Yplus)
        {

            double x = control.Current.BoundingRectangle.Location.X + Xplus;
            double y = control.Current.BoundingRectangle.Location.Y + Yplus;
            MouseClick m = new MouseClick();
            m.DoMouseClick((uint)x, (uint)y);
        }

        public string buildfilename(string filename)
        {

            string fname = "";
            string year = Convert.ToString(DateTime.Now.Year);
            string month = Convert.ToString(DateTime.Now.Month).PadLeft(2, '0');
            string days = Convert.ToString(DateTime.Now.Day).PadLeft(2, '0');

            fname = filename + month + days + year;
            return fname;
        }

        public string buildpath(string partialpath)
        {

            string fullpath = "";
            string year = Convert.ToString(DateTime.Now.Year);
            string month = Convert.ToString(DateTime.Now.Month).PadLeft(2, '0');
            string monthname = Convert.ToString(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month));

            fullpath = partialpath + @"\" + year + @"\" + month + "-" + monthname.Substring(0, 3) + @"\";

            return fullpath;
        }

       public void createPath(string path) {

            bool ifExists = System.IO.Directory.Exists(path);

            if (!ifExists)
            {
                System.IO.Directory.CreateDirectory(path);
            }
        
        }

       public TabPage getTabPage(string name) {

           TabPage tabitem = mainwind.Get<TabPage>(SearchCriteria.ByText(name));
           return tabitem;
       }


       public void LockAdminLoops(int loop) {

           for (int i = 0; i < loop; i++)
           {
               KeyStrokes key = new KeyStrokes();
               key.sendvalue("+{DOWN}");
               

           }
         
       
       }

       public DateTime previousFriday() 
       { 
       int fridays=0;
       DateTime prevFriday=DateTime.Today;
       while (fridays < 2)
       {

           if (prevFriday.DayOfWeek.ToString() == "Friday")
           {

               fridays += 1;
               if (fridays < 2)
               {

                   prevFriday = prevFriday.AddDays(-1);
               }
           }
           else {

               prevFriday = prevFriday.AddDays(-1);
           
           }
       
       
       }

       return prevFriday;
       }

       public void clickMenuItem(MenuBar menu, string name) {

           menu.MenuItem(name).Click();
       
       }

       public ListView getDataGrid(Window wind) 
       {

           ListView table = wind.Get<ListView>(SearchCriteria.ByClassName("TMListView"));

          return table;
       }

       public bool CheckBIWeekly()
       {
           bool answer=true;
           DateTime date = Convert.ToDateTime(ConfigurationManager.AppSettings["BIWeeklyStartDate"]);

           int numberofdays = Convert.ToInt32(DateTime.Today.Subtract(date).TotalDays);

           if ((numberofdays % 2) == 0 && DateTime.Today.DayOfWeek.ToString() == ConfigurationManager.AppSettings["BIWeeklyDay"])
           {
               answer = true;
              

           }
           else {

               answer = false;
           }
          
           return answer;
       }

       public bool SplitWeek(int days) {

           bool split = false;
           DateTime lastdaterun = DateTime.Now.AddDays(days);
           if (lastdaterun.Month != DateTime.Now.Month)
           {

               split = true;

           }
           else
           {

               split = false;
           }

           return split;
       
       }

       public bool checkSundays()
       {
           bool sunday = false;
           if (DateTime.Today.DayOfWeek.ToString() == "Sunday")
           {

               sunday = true;

           }
           return sunday;
       }


       public bool checkSaturdays()
       {
           bool sat = false;
           if (DateTime.Today.DayOfWeek.ToString() == "Saturday")
           {

               sat = true;

           }
           return sat;
       }

       public bool checkWeekday()
       {
           bool week = true;

           switch(DateTime.Today.DayOfWeek.ToString()){
                  
               case "Saturday":

                   week = false;
                   break;

               case "Sunday":

                    week = false;
                    break;
           
           
           }
           
           return week;
       }

       public bool checkEOM()
       {
           bool eom = false;
           if (DateTime.Today.Day == DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month) || DateTime.Today.Day == 1)
           {

               eom = true;

           }
           return eom;
       }

       public bool checkMidM()
       {
           bool midm = false;
           if (DateTime.Today.Day == 15)
           {

               midm = true;

           }
           return midm;
       }

       public Panel getPane(int i, Window wind)
       {

           Panel pane = wind.Get<Panel>(SearchCriteria.ByText("").AndIndex(i));
           app.WaitWhileBusy();
           return pane;

       }


       public void isProcessIddle(string processname) {
           try
           {
               int count = 0;
               Process[] pname = Process.GetProcessesByName(processname);
               foreach (Process myprocess in pname)
               {

                   while (count < Convert.ToInt32(ConfigurationManager.AppSettings["isIddleWaitTime"]))
                   {

                       TimeSpan begin_cpu_time = myprocess.TotalProcessorTime;
                       System.Threading.Thread.Sleep(3000);
                       myprocess.Refresh();
                       TimeSpan end_cpu_time = myprocess.TotalProcessorTime;
                       if (end_cpu_time - begin_cpu_time == TimeSpan.FromMilliseconds(0))
                       {
                           count += 1;
                       }
                       else
                       {

                           count = 0;
                       }

                   }
               }
           }
           catch (Exception ex)
           {

               logfile.writefile(ex.Message.ToString(), ConfigurationManager.AppSettings[@"LogFile"]);
              

           }
       
       }

     
     
       
        ////*********************METHODS END********************************
    }
}
