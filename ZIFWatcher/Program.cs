using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ZIFWatcher
{
    static class Program
    {
        private static Mutex Mut_Ex = null;
        private static string APPname = "ZIF.Watcher";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Mut_Ex = new Mutex(true, APPname, out bool rulareNoua);
            if (rulareNoua == false) { return; }

            string appFN = Path.GetFileNameWithoutExtension(Process.GetCurrentProcess().MainModule.FileName);
            if (!appFN.Equals(APPname, StringComparison.InvariantCultureIgnoreCase))
            {
                MessageBox.Show($"The APP's filename is '{appFN}', but it should be '{APPname}'. This is required for other components to work properly. Please rename the file."
                                + Environment.NewLine + Environment.NewLine + "The app will close, sorry!", APPname, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            frmMain fereastra_Start = new frmMain();
            fereastra_Start.Show();

            Application.Run();
        }
    }
}
