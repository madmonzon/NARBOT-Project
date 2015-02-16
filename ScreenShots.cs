using System;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;

namespace NARBOT
{
    class ScreenShots
    {
        public void printScreen(string filename)
        {

            Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(bitmap as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
            bitmap.Save(ConfigurationManager.AppSettings[@"PrintScreenPath"].ToString() + filename+".png");

        }
    }
}
