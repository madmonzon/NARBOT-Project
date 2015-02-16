using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System.Configuration;

namespace NARBOT
{
    class PDFtoXML
    {
        WriteLog logfile = new WriteLog();
        string log = "";

        public void ConvertPDFtoXML(string InputFileName)
        {
            string[] lines, tags, ts = { "", "  ", "    ", "      ", "        ", "          " }; // How much to indent the tags in the output file
            int StartIndex, EndIndex, IndentLevel = 1, NumStatements = 0;
            decimal totalAmount = 0;
            bool delayIndent = false;
            string OutputFileName = InputFileName.Replace(".pdf", ".xml");
            PDFTextStripper stripper=new PDFTextStripper();
            PDDocument doc=new PDDocument();
            try
            {
                doc = PDDocument.load(InputFileName);
                stripper = new PDFTextStripper();
            }
            catch (Exception ex)
            {
                log = ex.Message.ToString();
                logfile.writefile(log, ConfigurationManager.AppSettings[@"LogFile"]);
               
            }
            string AllPDFtext = stripper.getText(doc); // Returns all text from PDF file
            lines = AllPDFtext.Replace("\r", "").Split('\n'); // Work with line units and remove any CRs
            using (StreamWriter sw = new StreamWriter(OutputFileName))
            {
                sw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");
                foreach (string line in lines)
                {
                    if (line.StartsWith("<") || line.EndsWith(">")) // Found tag(s) on line
                    {
                        StartIndex = line.IndexOf('<');
                        EndIndex = line.LastIndexOf('>');
                        // If multiple tags per line, break into singles. Some multi-tags have a spece between them, others abut
                        tags = line.Substring(StartIndex, (EndIndex - StartIndex) + 1).Replace("> <", ">~<").Replace("><", ">~<").Split('~');
                        foreach (string tag in tags)
                        {
                            if (tag == "<Statement>") //Count these
                                ++NumStatements;
                            else if (tag.StartsWith("<BillAmount>")) //Accumulate these
                                totalAmount += Convert.ToDecimal(getValueFromTag(tag));
                            else if (tag.Contains("ControlTotals>")) //Don't print these
                                continue;
                            else if (tag.Contains("<TotalBilled>")) //Compare our totals to these
                            {
                                totalAmount -= Convert.ToDecimal(getValueFromTag(tag));
                               
                            }
                            else if (tag.Contains("<TotalBills>")) //Compare our totals to these
                            {
                                NumStatements -= Convert.ToInt32(getValueFromTag(tag).Replace(",", ""));
                              
                            }
                            if (tag.LastIndexOf('<') == 0) //single tag on a line (only one '<')
                            {
                                if (tag[1] == '/' && IndentLevel > 0) // Single "close" tag
                                    --IndentLevel;
                                else // Single "open" tag
                                    delayIndent = true;
                            }
                            sw.Write(ts[IndentLevel]);
                            sw.WriteLine(tag.Replace("&", "&amp;"));
                            if (delayIndent) //Makes the output look even
                            {
                                delayIndent = false;
                                if (IndentLevel < ts.Length - 1) //Make sure we don't overstep the array
                                    ++IndentLevel;
                            }
                        } 
                    } 
                } 
            } 
            
        } 

        static string getValueFromTag(string tag)
        {
            int StartIndex, EndIndex;
            StartIndex = tag.IndexOf('>') + 1;
            EndIndex = tag.Substring(StartIndex).IndexOf('<');
            return tag.Substring(StartIndex, EndIndex);
        }
    }
}
