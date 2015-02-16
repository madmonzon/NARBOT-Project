using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Configuration;
using System.IO;

namespace NARBOT
{
    class FTPSend
    {

        WriteLog logfile = new WriteLog();
        string log = "";

        public void ftpcopy(string fileurl, string filename)
        {


            FtpWebRequest ftpClient = (FtpWebRequest)FtpWebRequest.Create(ConfigurationManager.AppSettings["FTPServer"] + filename);
            ftpClient.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["ftpuname"], ConfigurationManager.AppSettings["ftppass"]);
            ftpClient.Method = System.Net.WebRequestMethods.Ftp.UploadFile;

            byte[] fileContents = File.ReadAllBytes(@fileurl);

            ftpClient.ContentLength = fileContents.Length;

            Stream requestStream = ftpClient.GetRequestStream();
            requestStream.Write(fileContents, 0, fileContents.Length);
            requestStream.Close();

            FtpWebResponse response = (FtpWebResponse)ftpClient.GetResponse();

            log = response.StatusCode.ToString();

            logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);

            response.Close();
        }

    }
}
