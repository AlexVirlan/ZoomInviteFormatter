using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ZoomInviteFormatter.Clase;

namespace ZoomInviteFormatter
{
    public partial class frm_Start : Form
    {
        public frm_Start(bool startedByWatcher = false)
        {
            InitializeComponent();
            cls_Variabile.StartedByWatcher = startedByWatcher;
        }

        #region Variabile
        int loadIndex = 0;
        bool loadReverse = false;
        #endregion

        #region Functii Auxiliare
        private void DORMI(int Interval)
        {
            System.Diagnostics.Stopwatch Cronometru = new System.Diagnostics.Stopwatch();
            Cronometru.Start();
            while (Cronometru.ElapsedMilliseconds < Interval)
            {
                Application.DoEvents();
            }
            Cronometru.Stop();
        }
        #endregion

        private void frm_Start_Load(object sender, EventArgs e)
        {
            lbl_StartedByWatcher.Visible = cls_Variabile.StartedByWatcher;
        }

        private void frm_Start_Shown(object sender, EventArgs e)
        {
            while (this.Opacity <= 0.92)
            {
                this.Opacity += 0.02;
                DORMI(12);
            }
            DORMI(260);
            lbl_oom.Visible = true;
            DORMI(460);
            lbl_nvite.Visible = true;
            DORMI(460);
            lbl_ormatter.Visible = true;
            DORMI(260);
            tmr_Loading.Interval = 170;
            tmr_Loading.Start();
            Thread tLoad = new Thread(new ThreadStart(INCARCARE));
            tLoad.Start();
        }

        private void INCARCARE()
        {
            cls_Variabile.CaleDateAPP = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\AvA.Soft\ZIF";
            if (Directory.Exists(cls_Variabile.CaleDateAPP) == false)
            {
                Directory.CreateDirectory(cls_Variabile.CaleDateAPP);
            }
            if (File.Exists(cls_Variabile.CaleDateAPP + "\\APP.set") == false)
            {
                cls_Functii.ReseteazaPreseturi();
                RaspunsFunctie rF = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rF.Eroare)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        this.TopMost = false;
                        MessageBox.Show(this, "Error - saving default settings:" + Environment.NewLine +
                                        rF.Mesaj + Environment.NewLine + Environment.NewLine +
                                        "The app will close. :(", "Z I F", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                        return;
                    }));
                }
            }
            else
            {
                RaspunsFunctie rF = cls_Setari.Incarca(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rF.Eroare)
                {
                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        this.TopMost = false;
                        MessageBox.Show(this, "Error - loading settings:" + Environment.NewLine +
                                        rF.Mesaj + Environment.NewLine + Environment.NewLine +
                                        "The app will close. :(", "Z I F", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                    Application.Exit();
                    return;
                }
            }
            //
            if (Setari.AppRuns == 0) { cls_Variabile.PrimaRulare = true; }
            Setari.AppRuns++;
            cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
            //
            DORMI(3300);
            this.Invoke(new MethodInvoker(delegate ()
            {
                frm_Main mainForm = new frm_Main(this);
                mainForm.Show();
            }));
        }

        private void tmr_Loading_Tick(object sender, EventArgs e)
        {
            if (loadReverse) { loadIndex--; }
            else { loadIndex++; }

            if (loadIndex == 6) { loadReverse = true; }
            if (loadIndex == 0) { loadReverse = false; }

            string dots = "• ".Repeat(loadIndex);
            lbl_Load.Text = dots + "  l o a d i n g   " + dots;
        }
    }

    public static class StringExtensions
    {
        public static string Repeat(this string s, int n) => new StringBuilder(s.Length * n).Insert(0, s, n).ToString();
    }
}
