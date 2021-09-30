using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (var hiddenForm = new Form())
            using (var mag = new Karna.Magnification.Magnifier(hiddenForm) { Magnification = 1f })
            {
                Application.Run();
            }
        }
    }
}
