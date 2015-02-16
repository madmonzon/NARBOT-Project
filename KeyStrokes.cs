using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace NARBOT
{
    class KeyStrokes
    {

        public void sendspace() {

            SendKeys.SendWait("( )");
        
        }

        public void sendvalue(string val)
        {

            SendKeys.SendWait(val);

        }
    }
}
