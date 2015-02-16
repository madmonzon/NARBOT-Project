using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Web;
using System.Configuration;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace NARBOT
{
    class CrystalReports
    {
        public string memoinvpdfreportandpath;
        Core bot = new Core();
        public void CR(string distcode, string letter, string reportname, bool multivalue)
        {

            ReportDocument rpt = new ReportDocument();
            rpt.Load(ConfigurationManager.AppSettings[@"StatementReportPath"]);
           
            DataSourceConnections conn = rpt.DataSourceConnections;
            
            for (int cnt = 0; cnt < conn.Count; cnt++)
            {
                conn[cnt].SetLogon(ConfigurationManager.AppSettings["UsernameStat"], ConfigurationManager.AppSettings["PasswordStat"]); //Might have to check connectionName to determine login.
                
            }

            
            rpt.SetParameterValue(0, "*");
         
            if (multivalue)
            {
                rpt.SetParameterValue(1, new string[] { "Company", "Statement Copy" });
            }
            else {

                rpt.SetParameterValue(1, "Company");
            }
            rpt.SetParameterValue(2, distcode);
            rpt.SetParameterValue(3,letter);
            rpt.SetParameterValue(4, "ALL");

           

          
            try // Export
            {
                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                bot.createPath(ConfigurationManager.AppSettings[@"SaveStatReportPath"]);
                CrDiskFileDestinationOptions.DiskFileName = ConfigurationManager.AppSettings[@"SaveStatReportPath"] + reportname + ".pdf";
                CrExportOptions = rpt.ExportOptions;
              
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.ExportDestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.ExportFormatOptions = CrFormatTypeOptions;
               
                rpt.Export();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            rpt.Close();
            rpt.Dispose();

        }

        public void CRMemoInv(string reportname)
        {

            ReportDocument rpt = new ReportDocument();
            rpt.Load(ConfigurationManager.AppSettings[@"MemoBillInvPath"]);

            DataSourceConnections conn = rpt.DataSourceConnections;

            for (int cnt = 0; cnt < conn.Count; cnt++)
            {
                conn[cnt].SetLogon(ConfigurationManager.AppSettings["UsernameStat"], ConfigurationManager.AppSettings["PasswordStat"]); //Might have to check connectionName to determine login.

            }


            rpt.SetParameterValue(0, "All");

            try // Export
            {
                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                string fullsavepath=bot.buildpath(ConfigurationManager.AppSettings[@"MemoBillPath"]);
                memoinvpdfreportandpath = fullsavepath + ConfigurationManager.AppSettings[@"MemoInvDir"] + @"\" + reportname + ".pdf";
                CrDiskFileDestinationOptions.DiskFileName = memoinvpdfreportandpath;
                CrExportOptions = rpt.ExportOptions;

                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                CrExportOptions.ExportDestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.ExportFormatOptions = CrFormatTypeOptions;

                rpt.Export();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            rpt.Close();
            rpt.Dispose();

        }
    }
}
