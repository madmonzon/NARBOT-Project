using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace NARBOT
{
    class WriteLog
    {
        public void writefile(string text, string filepath)
        {
         
                writetext(text, filepath);
           

        }
        private void writetext(string text, string filepath)
        {


            string filename = filepath + ConfigurationManager.AppSettings["LogFileName"];

            System.IO.StreamWriter objwriter;

            objwriter = new System.IO.StreamWriter(filename, true);
            objwriter.WriteLine(text);

            objwriter.Close();




        }
    }
}
