using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows;
using System.Configuration;


namespace NARBOT
{
    class DefaultPrinter
    {
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);

        public void setprinter() {


            string pname = ConfigurationManager.AppSettings["DefaultPrinter"];
           DefaultPrinter.SetDefaultPrinter(pname);
       
        }
    }
}
