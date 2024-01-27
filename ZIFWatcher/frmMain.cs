using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ZIFWatcher
{
    public partial class frmMain : Form
    {
        #region Vars
        private bool _zifDejaPornit = false;
        #endregion

        #region Functii Auxiliare
        public bool IsProcessRunning(string processName)
        {
            return Process.GetProcessesByName(processName).Count() > 0;
        }
        //
        //public void KillProcess(string processName)
        //{
        //    Process.GetProcessesByName(processName).ToList().ForEach(p => p.Kill());
        //}
        //public int KillProcess(string processName)
        //{
        //    Process[] processes = Process.GetProcessesByName(processName);
        //    foreach (Process P in processes)
        //    {
        //        P.Kill();
        //    }
        //    return processes.Count();
        //}
        #endregion

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(2600);
            tmr_Watcher.Interval = 2000;
            tmr_Watcher.Start();
        }

        private void tmr_Watcher_Tick(object sender, EventArgs e)
        {
            if (File.Exists(".watcher") && IsProcessRunning("Zoom") && !IsProcessRunning("ZoomInviteFormatter") && !_zifDejaPornit)
            {
                _zifDejaPornit = true;
                Process proccc = Process.Start("ZoomInviteFormatter", "startedByWatcher");
            }
            if (_zifDejaPornit && !IsProcessRunning("Zoom"))
            {
                _zifDejaPornit = false;
            }
        }

        private void RunZIF()
        {
            Process P = new Process();
            ProcessStartInfo pSi = new ProcessStartInfo();
            pSi.FileName = "ZoomInviteFormatter";
            //pSi.WorkingDirectory = "";
            pSi.Arguments = "startedByWatcher";
            //pSi.CreateNoWindow = false;
            //pSi.WindowStyle = ProcessWindowStyle.Normal;
            //pSi.UseShellExecute = false;
            P.StartInfo = pSi;
            P.Start();
        }
    }
}
