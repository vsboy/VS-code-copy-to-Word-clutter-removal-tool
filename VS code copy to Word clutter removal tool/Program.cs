using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace VS_code_copy_to_Word_clutter_removal_tool
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
