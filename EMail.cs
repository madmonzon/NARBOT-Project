using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using System.Xml;
using System.Xml.Linq;

namespace NARBOT
{
    class EMail
    {

        WriteLog logfile = new WriteLog();

        public void email_send(string body, string attpath)
        {
            string log = "";
            int tries = 0;
            while (tries < Convert.ToInt32(ConfigurationManager.AppSettings["EmailTries"]))
            {
                try
                {
                    MailMessage mail = new MailMessage();
                    SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SMTP"]);
                    mail.From = new MailAddress(getelement("from"));
                    mail.To.Add(getelement("to"));
                    mail.Subject = getelement("subject") + "-" + body;
                    mail.Body = body + " completed.";

                    if (attpath != "")
                    {
                        System.Net.Mail.Attachment attachment;
                        attachment = new System.Net.Mail.Attachment(attpath);
                        mail.Attachments.Add(attachment);

                    }
                    SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["EmailPort"]);

                    SmtpServer.Send(mail);
                    tries = Convert.ToInt32(ConfigurationManager.AppSettings["EmailTries"]);

                }
                catch (Exception ex)
                {

                    log = ex.Message.ToString();
                    logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);

                    tries += 1;

                }
            }

        }

        public string getelement(string nodename)
        {
            string str = "";
            XmlDocument doc = new XmlDocument();
            doc.Load(ConfigurationManager.AppSettings[@"EMailXMLDir"]);
            XmlElement root = doc.DocumentElement;
            XmlNodeList lst = root.GetElementsByTagName(nodename);
            foreach (System.Xml.XmlNode n in lst)
            {

                int nodecounter = 0;
                nodecounter = n.ChildNodes.Count;
                while (nodecounter > 0)
                {

                    XmlNode name = n.FirstChild;
                    str = name.InnerText;
                    n.RemoveChild(name);
                    nodecounter--;
                }
            }
            return str;
        }

    }
}
