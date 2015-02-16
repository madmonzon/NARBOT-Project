using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows;

namespace NARBOT
{
    class MouseClick
    {
        [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
   public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

   private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
   private const int MOUSEEVENTF_LEFTUP = 0x0004;
   private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
   private const int MOUSEEVENTF_RIGHTUP = 0x10;


   public void DoMouseClick(uint X, uint Y)
   {
      //Call the imported function with the cursor's current position
       Cursor.Position = new System.Drawing.Point((int)X, (int)Y);
       mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
    
   }

    }
}
        