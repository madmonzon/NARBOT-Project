using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Configuration;

namespace NARBOT
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            try
            {
                Core bot = new Core();
                bot.createPath(ConfigurationManager.AppSettings[@"PrintScreenPath"]);
                bot.createPath(ConfigurationManager.AppSettings[@"LogFile"]);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString(), "Directory Creation Error", MessageBoxButtons.OK);

            }

            try
            {
                DefaultPrinter printer = new DefaultPrinter();
                printer.setprinter();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString(), "Default Printer Error", MessageBoxButtons.OK);


            }

             try
            {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            }
             catch (Exception ex)
             {

                 MessageBox.Show(ex.Message.ToString(), "BOT Start Error", MessageBoxButtons.OK);

             }
        }
    }
}
