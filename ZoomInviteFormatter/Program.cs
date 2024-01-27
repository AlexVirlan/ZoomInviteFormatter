using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace ZoomInviteFormatter
{
    static class Program
    {
        private static Mutex Mut_Ex = null;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Mut_Ex = new Mutex(true, "Z.I.F.", out bool rulareNoua);
            if (rulareNoua == false) { return; }

            string appFN = Path.GetFileNameWithoutExtension(Process.GetCurrentProcess().MainModule.FileName);
            if (!appFN.Equals("ZoomInviteFormatter", StringComparison.InvariantCultureIgnoreCase))
            {
                MessageBox.Show($"The APP's filename is '{appFN}', but it should be 'ZoomInviteFormatter'. This is required for other components to work properly. Please rename the file."
                                + Environment.NewLine + Environment.NewLine + "The app will close, sorry!", "ZIF (Zoom Invite Formatter)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string[] args = Environment.GetCommandLineArgs().Skip(1).ToArray();
            bool byWatcher = args.Count() > 0 && args[0].Equals("startedByWatcher", StringComparison.InvariantCultureIgnoreCase);

            frm_Start fereastra_Start = new frm_Start(byWatcher);
            fereastra_Start.Show();
            Application.Run();

            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new frm_Main());
        }
    }
}
