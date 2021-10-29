using System;
using System.Collections.Generic;
using System.Drawing;
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

            Karna.Magnification.Fork.Magnifier.DoStuff += Callback;

            using (var hiddenForm = new Form())
            using (var mag = new Karna.Magnification.Fork.Magnifier(hiddenForm) { Magnification = 1f })
            {
                hiddenForm.Show();
                Application.Run();
            }
        }

        static void Callback(Bitmap b)
        {
            Console.WriteLine(b);
        }
    }
}
