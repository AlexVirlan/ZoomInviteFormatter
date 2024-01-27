using Microsoft.VisualBasic;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using ZoomInviteFormatter.Clase;

namespace ZoomInviteFormatter
{
    public partial class frm_Main : Form
    {
        #region Variabile
        private readonly string _TitluAPP = "ZIF (Zoom Invite Formatter)";
        private Form _StartForm;
        private ExportImport _exportImport = new ExportImport();
        private string _dateFisierImport = "";
        private bool _preseturiModificate = false;
        Font _lblActiv;
        Font _lblInactiv;
        private static bool _isCtrlDown = false;
        private GlobalKeyboardHook _globalKeyboardHook;
        static string _NL = Environment.NewLine;
        #endregion

        public frm_Main(Form fereastraStart)
        {
            InitializeComponent();
            _StartForm = fereastraStart;
        }

        #region Functii Auxiliare
        private void DORMI(int Interval)
        {
            Stopwatch Cronometru = new Stopwatch();
            Cronometru.Start();
            while (Cronometru.ElapsedMilliseconds < Interval)
            {
                Application.DoEvents();
            }
            Cronometru.Stop();
        }
        #endregion

        private void frm_Main_Load(object sender, EventArgs e)
        {
            if (Setari.Preseturi.Count > 0)
            {
                cls_Functii.IncarcaPreseturi(cmb_Preseturi, lbl_Presets);
                if (Setari.RememberLastTemplate && Setari.LastTemplateIndex < Setari.Preseturi.Count) 
                { cmb_Preseturi.SelectedIndex = Setari.LastTemplateIndex; }
            }
            lbl_Runs.Text = $"Hello {Environment.UserName}! You opened this app {Setari.AppRuns} times.   |   AvA.Soft - We Make IT Happen!";
            rtb_Mesaj_TextChanged(null, null);
            txt_Preview.Location = pnl_Mesaj.Location;
            pnl_SET.Location = pnl_MAIN.Location;
            cls_Functii.FillComboWithWeekDays(cmb_Days);
            cmb_Days.SelectedIndex = 0;
            IncarcaSetari();
            cls_Variabile.VersiuneApp = $"{Assembly.GetExecutingAssembly().GetName().Version.Major}.{Assembly.GetExecutingAssembly().GetName().Version.Minor}";
            this.Text = $"ZIF (Zoom Invite Formatter) v.{cls_Variabile.VersiuneApp} - AvA.Soft";
            lbl_CurrVersion.Text = $"Version: {cls_Variabile.VersiuneApp}";
            lbl_Email.Text = $"Email: {cls_Variabile.ContactMail}";
            lbl_Telegram.Text = $"Telegram: @{cls_Variabile.ContactTelegram}";
            lbl_WebSite.Text = $"Web: {cls_Variabile.ContactWebsite}";
            cmb_WhatToExport.SelectedIndex = cmb_ImportMode.SelectedIndex = 0;
            _lblActiv = lbl_Mexport.Font;
            _lblInactiv = lbl_Mimport.Font;
            if (cls_Variabile.StartedByWatcher) { this.Text += " - [auto-started]"; }
            SetupKeyboardHooks();
        }

        private void frm_Main_Shown(object sender, EventArgs e)
        {
            while (_StartForm.Opacity >= 0.02)
            {
                _StartForm.Opacity -= 0.02;
                DORMI(12);
            }
            _StartForm.Close();
            _StartForm.Dispose();
            while (this.Opacity <= 0.92)
            {
                this.Opacity += 0.02;
                DORMI(12);
            }
            this.Opacity = Setari.Trans;
            if (cls_Variabile.PrimaRulare)
            {
                MessageBox.Show(this, $"Hi {Environment.UserName}! Welcome to 'ZIF'! :)" + _NL +
                                "Before we can start, please set your meetings hours." + _NL + _NL +
                                "Thank you for using this app!",
                                _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                ArataSETARI();
            }
            else
            {
                if (Setari.Preseturi.Count == 0)
                {
                    MessageBox.Show(this, "You don't have any templates. :(" + _NL +
                                    "Feel free to add a new one or reset to the default templates.",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (Setari.Intruniri.Count > 0) { lbl_GetTime_Click(null, null); }
                else
                {
                    MessageBox.Show(this, "You don't have any meetings. :(" + _NL +
                                    "You can add them so you can use the {time} tag.",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            tmr_AutoMode.Enabled = true;
            tmr_AutoMode.Start();
        }

        #region KB Hook
        public void SetupKeyboardHooks()
        {
            //_globalKeyboardHook = new GlobalKeyboardHook();
            _globalKeyboardHook = new GlobalKeyboardHook(new Keys[] { Keys.LControlKey, Keys.RControlKey });
            _globalKeyboardHook.KeyboardPressed += OnKeyPressed;
        }
        private void OnKeyPressed(object sender, GlobalKeyboardHookEventArgs e)
        {
            if (e.KeyboardData.Key == Keys.LControlKey || e.KeyboardData.Key == Keys.RControlKey)
            {
                if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.KeyDown)
                {
                    _isCtrlDown = true;
                    btn_Copy.Text = "C O P Y    T H E    T E X T   (and exit ZIF)";
                    btn_Copy.BackColor = Color.FromArgb(46, 46, 64);
                }
                else if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.KeyUp)
                {
                    _isCtrlDown = false;
                    btn_Copy.Text = "C O P Y    T H E    T E X T";
                    btn_Copy.BackColor = Color.FromArgb(46, 64, 46);
                }
            }

            //Debug.WriteLine(e.KeyboardData.VirtualCode);
            //if (e.KeyboardData.VirtualCode != GlobalKeyboardHook.VkSnapshot)
            //    return;
            //if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.KeyDown)
            //{
            //    MessageBox.Show("Print Screen");
            //    e.Handled = true;
            //}
            //if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.SysKeyDown &&
            //    e.KeyboardData.Flags == GlobalKeyboardHook.LlkhfAltdown)
            //{
            //    //MessageBox.Show("Alt + Print Screen");
            //    //e.Handled = true;
            //}
        }
        #endregion

        private void cmb_Preseturi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_Preseturi.SelectedItem is null)
            {
                rtb_Mesaj.Text = "";
                return;
            }
            rtb_Mesaj.Text = Setari.Preseturi[cmb_Preseturi.SelectedItem.ToString()];
        }

        private void btn_Edit_Click(object sender, EventArgs e)
        {
            if (cmb_Preseturi.SelectedIndex < 0)
            {
                MessageBox.Show(this, "You don't have any templates. :(" + _NL +
                                "Feel free to add a new one or reset to the default templates.",
                                _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            rtb_Mesaj.ReadOnly = false;
            rtb_Mesaj.ForeColor = Color.White;
            btn_Edit.Enabled = false;
            btn_Save.Enabled = btn_Cancel.Enabled = btn_Preview.Enabled = true;
            pnl_Presets.Enabled = false;
            pnl_Mesaj.Visible = true;
            txt_Preview.Visible = false;
            btn_Copy.Enabled = false;
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            try
            {
                Setari.Preseturi[cmb_Preseturi.SelectedItem.ToString()] = rtb_Mesaj.Text;
                RaspunsFunctie rf = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rf.Eroare)
                {
                    MessageBox.Show(this, "Error saving settings:" + _NL +
                                    rf.Mesaj + _NL + _NL +
                                    "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                else
                {
                    MessageBox.Show(this, "The text was saved successfully. :)",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error deleting the text:" + _NL +
                                ex.Message + _NL + _NL +
                                "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            finally
            {
                rtb_Mesaj.ReadOnly = true;
                rtb_Mesaj.ForeColor = Color.Gray;
                btn_Edit.Enabled = true;
                btn_Save.Enabled = btn_Cancel.Enabled = btn_Preview.Enabled = false;
                pnl_Presets.Enabled = true;
                pnl_Mesaj.Visible = false;
                txt_Preview.Visible = true;
                btn_Copy.Enabled = true;
            }
        }

        private void btn_Delete_Click(object sender, EventArgs e)
        {
            if (cmb_Preseturi.SelectedIndex < 0)
            {
                MessageBox.Show(this, "You don't have any templates. :(" + _NL +
                                "Feel free to add a new one or reset to the default templates.",
                                _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            cms_Delete.Show(btn_Delete, new Point(0, btn_Delete.Height));
        }

        private void btn_Rename_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmb_Preseturi.SelectedIndex < 0)
                {
                    MessageBox.Show(this, "You don't have any templates. :(" + _NL +
                                    "Feel free to add a new one or reset to the default templates.",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string renameResult = Interaction.InputBox("Enter a new name for this template", "Rename template - ZIF", cmb_Preseturi.SelectedItem.ToString()).Trim();
                if (!string.IsNullOrEmpty(renameResult))
                {
                    Setari.Preseturi.RenameKey(cmb_Preseturi.SelectedItem.ToString(), renameResult);
                    cls_Functii.IncarcaPreseturi(cmb_Preseturi, lbl_Presets, cmb_Preseturi.SelectedIndex);
                    RaspunsFunctie rf = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                    if (rf.Eroare)
                    {
                        MessageBox.Show(this, "Error saving settings:" + _NL +
                                        rf.Mesaj + _NL + _NL +
                                        "The app will cose. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                    else
                    {
                        MessageBox.Show(this, "The template was renamed successfully. :)",
                                        _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error renaming the template:" + _NL +
                                ex.Message + _NL + _NL +
                                "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult rez = MessageBox.Show(this, "Are you sure you want to reset to de default templates?",
                                                   "Reset templates - ZIF", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                                   defaultButton: MessageBoxDefaultButton.Button2);
                if (rez == DialogResult.No) { return; }
                cls_Functii.ReseteazaPreseturi();
                cls_Functii.IncarcaPreseturi(cmb_Preseturi, lbl_Presets);
                RaspunsFunctie rf = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rf.Eroare)
                {
                    MessageBox.Show(this, "Error saving settings:" + _NL +
                                    rf.Mesaj + _NL + _NL +
                                    "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                else
                {
                    MessageBox.Show(this, "The reset was successful. :)",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error reseting the templates:" + _NL +
                                ex.Message + _NL + _NL +
                                "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            
        }

        private void frm_Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (Setari.RememberLastTemplate)
                {
                    Setari.LastTemplateIndex = cmb_Preseturi.SelectedIndex;
                    cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                }
                _globalKeyboardHook?.Dispose();
                Application.Exit();
            }
        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            cms_Add.Show(btn_Add, new Point(0, btn_Add.Height));
        }

        private void rtb_Mesaj_TextChanged(object sender, EventArgs e)
        {
            ActualizeazaTextPreview();
            ColoreazaText();
            //Thread tLoad = new Thread(new ThreadStart(ColoreazaText));
            //tLoad.Start();
        }

        private void ActualizeazaTextPreview()
        {
            txt_Preview.Text = rtb_Mesaj.Text.Replace("\n", "\r\n")
                .Replace("{link}", txt_Link.Text)
                .Replace("{id}", txt_ID.Text)
                .Replace("{pw}", txt_PW.Text)
                .Replace("{tel1}", txt_Tel1.Text)
                .Replace("{tel2}", txt_Tel2.Text)
                .Replace("{time}", txt_Time.Text);
        }

        private void txt_Link_TextChanged(object sender, EventArgs e)
        {
            ActualizeazaTextPreview();
        }

        private void txt_ID_TextChanged(object sender, EventArgs e)
        {
            ActualizeazaTextPreview();
        }

        private void txt_PW_TextChanged(object sender, EventArgs e)
        {
            ActualizeazaTextPreview();
        }

        private void txt_Tel1_TextChanged(object sender, EventArgs e)
        {
            ActualizeazaTextPreview();
        }

        private void txt_Tel2_TextChanged(object sender, EventArgs e)
        {
            ActualizeazaTextPreview();
        }

        private void txt_Time_TextChanged(object sender, EventArgs e)
        {
            ActualizeazaTextPreview();
        }

        private void ColoreazaText()
        {
            int pos = rtb_Mesaj.SelectionStart;
            rtb_Mesaj.SelectAll();
            rtb_Mesaj.SelectionColor = Color.White;
            rtb_Mesaj.DeselectAll();
            rtb_Mesaj.SelectionStart = pos;
            rtb_Mesaj.SelectionLength = 0;

            rtb_Mesaj.HighlightText("{link}", txt_Link.ForeColor);
            rtb_Mesaj.HighlightText("{id}", txt_ID.ForeColor);
            rtb_Mesaj.HighlightText("{pw}", txt_PW.ForeColor);
            rtb_Mesaj.HighlightText("{tel1}", txt_Tel1.ForeColor);
            rtb_Mesaj.HighlightText("{tel2}", txt_Tel2.ForeColor);
            rtb_Mesaj.HighlightText("{time}", txt_Time.ForeColor);
        }

        private void rb_Auto_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_Auto.Checked)
            {
                lbl_ModeInfo.Text = "(Fast and easy: just go into your Zoom meeting, click on 'Invite', then click on 'Copy invitation'. That's all! I will do the rest. You will like this, trust me!)";
                tmr_AutoMode.Enabled = true;
                tmr_AutoMode.Start();
                //
                txt_Link.Enabled = txt_ID.Enabled = txt_PW.Enabled = txt_Tel1.Enabled = txt_Tel2.Enabled = false;
            }
        }

        private void rb_Manual_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_Manual.Checked)
            {
                lbl_ModeInfo.Text = "(You will have to manually enter all the info: the link, the ID, the passcode, both telephone numbers, and the time. Remember: none of them are mandatory.)";
                tmr_AutoMode.Enabled = false;
                tmr_AutoMode.Stop();
                //
                txt_Link.Enabled = txt_ID.Enabled = txt_PW.Enabled = txt_Tel1.Enabled = txt_Tel2.Enabled = true;
            }
        }

        private void btn_Copy_Click(object sender, EventArgs e)
        {
            List<string> taguri = new List<string>();
            if (rtb_Mesaj.Text.Contains("{link}") && string.IsNullOrEmpty(txt_Link.Text)) { taguri.Add("{link}"); }
            if (rtb_Mesaj.Text.Contains("{id}") && string.IsNullOrEmpty(txt_ID.Text)) { taguri.Add("{id}"); }
            if (rtb_Mesaj.Text.Contains("{pw}") && string.IsNullOrEmpty(txt_PW.Text)) { taguri.Add("{pw}"); }
            if (rtb_Mesaj.Text.Contains("{tel1}") && string.IsNullOrEmpty(txt_Tel1.Text)) { taguri.Add("{tel1}"); }
            if (rtb_Mesaj.Text.Contains("{tel2}") && string.IsNullOrEmpty(txt_Tel2.Text)) { taguri.Add("{tel2}"); }
            if (rtb_Mesaj.Text.Contains("{time}") && string.IsNullOrEmpty(txt_Time.Text)) { taguri.Add("{time}"); }
            if (taguri.Count > 0)
            {
                string ctrlInfo = _NL + _NL + "If so, please keep CTRL pressed while clicking 'Yes'.";
                DialogResult rez = MessageBox.Show(this, "It looks like you are using these tags in your template that don't have any data:" + _NL +
                                                   string.Join(", ", taguri) + _NL + _NL +
                                                   "Are you sure that you want to copy the text like this?" +
                                                   (_isCtrlDown ? ctrlInfo : ""),
                                                   "Some tags are missing data - ZIF", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                                   defaultButton: MessageBoxDefaultButton.Button2);
                if (rez == DialogResult.No) { return; }
            }

            Clipboard.Clear();
            Clipboard.SetText(txt_Preview.Text);

            if (_isCtrlDown)
            {
                _globalKeyboardHook?.Dispose();
                btn_Copy.Text = "Text copied! Exiting ZIF in 2 seconds...";
                DORMI(2000);
                Application.Exit();
                return;
            }

            btn_Copy.Text = "Done! Just paste the text wherever you want. :)";
            DORMI(4600);
            btn_Copy.Text = "C O P Y    T H E    T E X T";
        }

        private void tmr_AutoMode_Tick(object sender, EventArgs e)
        {
            string dateCP = Clipboard.GetText();
            if (string.IsNullOrEmpty(dateCP)) { return; }
            if (!dateCP.Contains("zoom.us/") && !dateCP.Contains("Meeting ID:") && !dateCP.Contains("Passcode:")) { return; }
            List<string> listDateCP = dateCP.Split('\n').ToList();
            if (listDateCP.Count() > 4) // TO DO - de verificat la viitoarele versiuni de Zoom
            {
                if (dateCP == txt_Preview.Text) { return; }
                for (int i = 0; i < listDateCP.Count; i++)
                {
                    if (string.IsNullOrEmpty(listDateCP[i].Replace("\r", ""))) { continue; }

                    if (listDateCP[i].StartsWith("https://") && listDateCP[i].Contains("zoom.us/")) { txt_Link.Text = listDateCP[i].Replace("\r", ""); }
                    if (listDateCP[i].StartsWith("Meeting ID: ")) { txt_ID.Text = listDateCP[i].Replace("Meeting ID: ", "").Replace("\r", ""); }
                    if (listDateCP[i].StartsWith("Passcode: ")) { txt_PW.Text = listDateCP[i].Replace("Passcode: ", "").Replace("\r", ""); }

                    if (listDateCP[i].Contains("#,,,,*") && listDateCP[i + 1].Contains("#,,,,*"))
                    {
                        txt_Tel1.Text = listDateCP[i].Split(' ')[0].Replace("\r", "");
                        txt_Tel2.Text = listDateCP[i + 1].Split(' ')[0].Replace("\r", "");
                    }
                }
            }
        }

        private void btn_SET_Click(object sender, EventArgs e)
        {
            if (pnl_MAIN.Visible)
            {
                btn_ShowExportImport.Visible = true;
                pnl_ExportImportData.Visible = false;
                IncarcaSetari();
                ArataSETARI();
            }
            else
            {
                if (_preseturiModificate)
                {
                    DialogResult rez = MessageBox.Show(this, "It looks like the presets were modified. You may (or may not) want to save them. :)" + _NL +
                                                       "If you go back now, these modifications will be lost!" + _NL + _NL +
                                                       "Are you sure that you want to exit the settings?",
                                                       "Presets modified - ZIF", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                                       defaultButton: MessageBoxDefaultButton.Button2);
                    if (rez == DialogResult.Yes)
                    {
                        _preseturiModificate = false;
                        ArataMAIN();
                    }
                }
                else
                {
                    ArataMAIN();
                }
            }
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            rtb_Mesaj.Text = Setari.Preseturi[cmb_Preseturi.SelectedItem.ToString()];
            rtb_Mesaj.ReadOnly = true;
            rtb_Mesaj.ForeColor = Color.Gray;
            btn_Edit.Enabled = true;
            btn_Save.Enabled = btn_Cancel.Enabled = btn_Preview.Enabled = false;
            pnl_Presets.Enabled = true;
            pnl_Mesaj.Visible = false;
            txt_Preview.Visible = true;
            btn_Copy.Enabled = true;
        }

        private void btn_AddIntrunire_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txt_Description.Text.Trim()))
            {
                MessageBox.Show(this, "The description is missing. :(" + _NL +
                                "Please add a description first.",
                                _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            dgv_Intruniri.Rows.Add(txt_Description.Text.Trim(), cmb_Days.SelectedItem, dtp_Ora.Value.ToString("HH:mm"));
            txt_Description.Text = "";
            cmb_Days.SelectedIndex = 0;
            dtp_Ora.Value = new DateTime(2017, 08, 26, 0, 0, 0);
            _preseturiModificate = true;
        }

        private void btn_DelSelIntrunire_Click(object sender, EventArgs e)
        {
            if (dgv_Intruniri.SelectedRows.Count == 0)
            {
                return;
            }
            dgv_Intruniri.Rows.RemoveAt(dgv_Intruniri.SelectedCells[0].RowIndex);
            _preseturiModificate = true;
        }

        private void btn_DelAllIntrunire_Click(object sender, EventArgs e)
        {
            DialogResult rez = MessageBox.Show(this, "Are you sure you want to delete all meetings?",
                                               "Delete all meetings - ZIF", MessageBoxButtons.YesNo, MessageBoxIcon.Question, 
                                               defaultButton: MessageBoxDefaultButton.Button2);
            if (rez == DialogResult.No) { return; }
            dgv_Intruniri.Rows.Clear();
            _preseturiModificate = true;
        }

        private void btn_ResetIntruniri_Click(object sender, EventArgs e)
        {
            dgv_Intruniri.Rows.Clear();
            foreach (Intrunire i in Setari.Intruniri)
            {
                dgv_Intruniri.Rows.Add(i.Nume, cls_Functii.GetDayFromDayNumber(i.Ziua), i.Ora);
            }
        }

        private void btn_SaveIntruniri_Click(object sender, EventArgs e)
        {
            try
            {
                Setari.Intruniri.Clear();
                foreach (DataGridViewRow rand in dgv_Intruniri.Rows)
                {
                    Setari.Intruniri.Add(new Intrunire()
                    {
                        Nume = rand.Cells["Description"].Value.ToString(),
                        Ziua = Convert.ToInt32(((ComboboxItem)rand.Cells["Day"].Value).Value),
                        Ora = rand.Cells["Time"].Value.ToString()
                    });
                }
                //
                Setari.RememberLastTemplate = chk_RememberLastTemplate.Checked;
                Setari.LastTemplateIndex = cmb_Preseturi.SelectedIndex;
                if (cmb_TimeMode.SelectedIndex == 0) { Setari.TimeMode = "next"; } else { Setari.TimeMode = "nearest"; }
                Setari.Trans = this.Opacity;
                Setari.KeepOnTop = chk_OnTop.Checked;
                if (Setari.KeepOnTop) { this.TopMost = true; } else { this.TopMost = false; }
                if (chk_OpenWithZoom.Checked) { using (File.Create(cls_Variabile.FisierWatcher)); }
                else { if (File.Exists(cls_Variabile.FisierWatcher)) { File.Delete(cls_Variabile.FisierWatcher); } }
                //
                RaspunsFunctie rf = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rf.Eroare)
                {
                    MessageBox.Show(this, "Error saving the meetings:" + _NL +
                                    rf.Mesaj + _NL + _NL +
                                    "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                else
                {
                    _preseturiModificate = false;
                    MessageBox.Show(this, "The meetings were saved successfully! :)",
                                           _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error saving the meetings:" + _NL +
                                ex.Message + _NL + _NL +
                                "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void lbl_GetTime_Click(object sender, EventArgs e)
        {
            List<string> ore = Setari.Intruniri.Where(w => w.Ziua == cls_Functii.GetDayNumberFromDayName(DateTime.Now.DayOfWeek.ToString()))
                                               .OrderBy(o => o.Ora)
                                               .Select(s => s.Ora)
                                               .ToList();
            if (ore.Count == 0) { goto NoData; }
            List<DateTime> date = new List<DateTime>();
            foreach (string ora in ore)
            {
                string[] oraSplit = ora.Split(':'); // split by current time separator from cultureInfo
                date.Add(new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, Convert.ToInt32(oraSplit[0]), Convert.ToInt32(oraSplit[1]), 0));
            }

            if (Setari.TimeMode == "next") { date.RemoveAll(r => r < DateTime.Now); }
            if (date.Count == 0) { goto NoData; }

            DateTime closest = date.OrderBy(t => Math.Abs((t - DateTime.Now).Ticks)).First();
            txt_Time.Text = closest.ToString("HH:mm");

            return;
            NoData:
            MessageBox.Show(this, "Bad news... :(" + _NL +
                            "I couldn't get the time for the next meeting for today." + _NL +
                            "Please go to settings and check your meetings list." + _NL + _NL +
                            "Or maybe there isn't any meetings today and you just love this app. :D",
                            _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void lbl_TimeModeHelp_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, "This setting affects how is the time calculated:" + _NL + _NL +
                            "Next meeting - the time of the NEXT meeting from the current day will be selected (if there is one in the meetings list)." +
                            _NL + _NL +
                            "Nearest meeting - the time of the NEAREST meeting from the current day will be selected (if there is one in the meetings list). Keep in mind that this could select a time from a meeting that is already passed (since it might be closer to the current time than the next meeting).",
                            _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void IncarcaSetari()
        {
            dgv_Intruniri.Rows.Clear();
            foreach (Intrunire i in Setari.Intruniri)
            {
                dgv_Intruniri.Rows.Add(i.Nume, cls_Functii.GetDayFromDayNumber(i.Ziua), i.Ora);
            }
            //
            chk_RememberLastTemplate.Checked = Setari.RememberLastTemplate;
            if (Setari.TimeMode == "next") { cmb_TimeMode.SelectedIndex = 0; } else { cmb_TimeMode.SelectedIndex = 1; }
            tb_Trans.Value = Convert.ToInt32(Setari.Trans / 0.01);
            lbl_Trans.Text = $"Opacity ({tb_Trans.Value}%):";
            chk_OnTop.Checked = Setari.KeepOnTop;
            if (Setari.KeepOnTop) { this.TopMost = true; }
            _preseturiModificate = false;
            chk_OpenWithZoom.Checked = File.Exists(cls_Variabile.FisierWatcher);
        }

        private void tb_Trans_Scroll(object sender, EventArgs e)
        {
            lbl_Trans.Text = $"Opacity ({tb_Trans.Value}%):";
            this.Opacity = tb_Trans.Value * 0.01;
        }

        private void ArataSETARI()
        {
            pnl_MAIN.Visible = false;
            pnl_SET.Visible = true;
            btn_SET.Text = "<   B a c k   <";
            lbl_Runs.Text = "ATTENTION! Going back without saving the settings, will result in the settings being lost!";
        }

        private void ArataMAIN()
        {
            pnl_MAIN.Visible = true;
            pnl_SET.Visible = false;
            btn_SET.Text = "Settings && Info";
            this.Opacity = Setari.Trans;
            lbl_Runs.Text = $"Hello {Environment.UserName}! You opened this app {Setari.AppRuns} times.   |   AvA.Soft - We Make IT Happen!";
            lbl_Runs.ForeColor = Color.Gray;
        }

        private void btn_Preview_MouseDown(object sender, MouseEventArgs e)
        {
            pnl_Mesaj.Visible = false;
            txt_Preview.Visible = true;
        }

        private void btn_Preview_MouseUp(object sender, MouseEventArgs e)
        {
            pnl_Mesaj.Visible = true;
            txt_Preview.Visible = false;
        }

        private void lbl_Email_Click(object sender, EventArgs e)
        {
            Process.Start($"mailto:{cls_Variabile.ContactMail}");
        }

        private void lbl_Telegram_Click(object sender, EventArgs e)
        {
            Process.Start("https://t.me/" + cls_Variabile.ContactTelegram);
        }

        private void lbl_WebSite_Click(object sender, EventArgs e)
        {
            //Process.Start(cls_Variabile.ContactWebsite);
        }

        private void chk_PwExport_CheckedChanged(object sender, EventArgs e)
        {
            txt_PwExport.Enabled = chk_ExportPwHide.Enabled = chk_PwExport.Checked;
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            if (chk_PwExport.Checked && string.IsNullOrEmpty(txt_PwExport.Text))
            {
                MessageBox.Show(this, "You checked the password option, please add one. :)",
                                        _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (chk_PwExport.Checked && txt_PwExport.TextLength < 4)
            {
                MessageBox.Show(this, "Sorry, the password must be at least 4 characters.",
                                      _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            switch (cmb_WhatToExport.SelectedIndex)
            {
                case 0: // Templates only
                    if (Setari.Preseturi.Count == 0)
                    {
                        MessageBox.Show(this, "Sorry, you don't have any templates in the app yet.",
                                      _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    break;
                case 1: // Meetings only
                    if (Setari.Intruniri.Count == 0)
                    {
                        MessageBox.Show(this, "Sorry, you don't have any meetings in the app yet.",
                                      _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    break;
                case 2: // Templates and meetings
                    if (Setari.Preseturi.Count == 0 && Setari.Intruniri.Count == 0)
                    {
                        MessageBox.Show(this, "Sorry, you don't have any templates or meetings in the app yet.",
                                      _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    break;
            }
            try
            {
                SaveFileD.Title = "Please select a location for your file";
                SaveFileD.Filter = "ZIF files (*.zif)|*.zif";
                SaveFileD.FilterIndex = 0;
                SaveFileD.FileName = $"ZIF - {cmb_WhatToExport.SelectedItem} ({DateTime.Now:dd.MM.yyyy})";
                if (SaveFileD.ShowDialog(this) == DialogResult.OK)
                {
                    ExportImport ei = new ExportImport();
                    switch (cmb_WhatToExport.SelectedIndex)
                    {
                        case 0: // Templates only
                            ei.Preseturi = Setari.Preseturi;
                            break;
                        case 1: // Meetings only
                            ei.Intruniri = Setari.Intruniri;
                            break;
                        case 2: // Templates and meetings
                            ei.Preseturi = Setari.Preseturi;
                            ei.Intruniri = Setari.Intruniri;
                            break;
                    }

                    string date = JsonConvert.SerializeObject(ei);
                    if (chk_PwExport.Checked) { date = cls_Cryptography.Encrypt(date, txt_PwExport.Text); }
                    File.WriteAllText(SaveFileD.FileName, date);

                    MessageBox.Show(this, "File saved successfully!" + _NL + _NL +
                                    "Location:" + _NL +
                                    SaveFileD.FileName,
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);

                    chk_PwExport.Checked = false;
                    txt_PwExport.Text = "";
                    chk_ExportPwHide.Checked = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error exporting data:" + _NL +
                                ex.Message + _NL + _NL +
                                "Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chk_ExportPwHide_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_ExportPwHide.Checked)
            { txt_PwExport.PasswordChar = '\0'; }
            else { txt_PwExport.PasswordChar = '•'; }
        }

        private void lbl_EncryptPw_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, $"Well, it's {DateTime.Now.Year}, sometimes we need a little privacy and security, don't you think?" + _NL +
                            "This option will encrypt your file so you can send it to your friends without worries, since it might contain sensitive data.",
                            _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void tsm_CurrentTemplate_Click(object sender, EventArgs e)
        {
            DialogResult rez = MessageBox.Show(this, "Are you sure you want to delete this template:" + _NL +
                                                   cmb_Preseturi.SelectedItem.ToString(),
                                                   "Delete template - ZIF", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                                   defaultButton: MessageBoxDefaultButton.Button2);
            if (rez == DialogResult.No) { return; }
            StergeTemplate(peToate: false);
        }

        private void tsm_AllTemplates_Click(object sender, EventArgs e)
        {
            DialogResult rez = MessageBox.Show(this, "Are you sure that you want to delete all the templates?",
                                                   "Delete all templates - ZIF", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                                   defaultButton: MessageBoxDefaultButton.Button2);
            if (rez == DialogResult.No) { return; }
            StergeTemplate(peToate: true);
        }

        private void StergeTemplate(bool peToate)
        {
            try
            {
                if (peToate) { Setari.Preseturi.Clear(); }
                else { Setari.Preseturi.Remove(cmb_Preseturi.SelectedItem.ToString()); }
                RaspunsFunctie rf = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rf.Eroare)
                {
                    MessageBox.Show(this, "Error saving settings:" + _NL +
                                    rf.Mesaj + _NL + _NL +
                                    "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                else
                {
                    if (peToate) { cmb_Preseturi.Items.Clear(); }
                    else { cmb_Preseturi.Items.RemoveAt(cmb_Preseturi.SelectedIndex); }
                    //
                    if (cmb_Preseturi.Items.Count == 0) { rtb_Mesaj.Text = ""; }
                    else { cmb_Preseturi.SelectedIndex = 0; }
                    //
                    MessageBox.Show(this, $"The {(peToate ? "templates were" : "template was")} deleted successfully. :)",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lbl_Presets.Text = $"Templates ({cmb_Preseturi.Items.Count}):";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error deleting {(peToate ? "all templates" : "the template")}:" + _NL +
                                ex.Message + _NL + _NL +
                                "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void tsm_EmptyTemplate_Click(object sender, EventArgs e)
        {
            AdaugaTemplate(cloneazaCurent: false);
        }

        private void tsm_CloneTemplate_Click(object sender, EventArgs e)
        {
            if (Setari.Preseturi.Count == 0)
            {
                MessageBox.Show(this, "There are no templates to clone from.",
                                _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            AdaugaTemplate(cloneazaCurent: true);
        }

        private void AdaugaTemplate(bool cloneazaCurent)
        {
            try
            {
                string addResult = Interaction.InputBox("Enter a name for the new template", "New template - ZIF", "").Trim();
                if (string.IsNullOrEmpty(addResult))
                {
                    MessageBox.Show(this, "The template name can't be empty.",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (Setari.Preseturi.ContainsKey(addResult))
                {
                    MessageBox.Show(this, "Ther is already a template with this name." + _NL +
                                    "The template name must be unique.",
                                    _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (cloneazaCurent) { Setari.Preseturi.Add(addResult, Setari.Preseturi[cmb_Preseturi.SelectedItem.ToString()]); }
                else { Setari.Preseturi.Add(addResult, ""); }
                RaspunsFunctie rf = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rf.Eroare)
                {
                    MessageBox.Show(this, "Error saving settings:" + _NL +
                                    rf.Mesaj + _NL + _NL +
                                    "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                else
                {
                    cmb_Preseturi.Items.Add(addResult);
                    cmb_Preseturi.SelectedIndex = cmb_Preseturi.Items.Count - 1;
                    MessageBox.Show(this, "The template was added successfully! :)",
                                           _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lbl_Presets.Text = $"Templates ({cmb_Preseturi.Items.Count}):";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error adding the template:" + _NL +
                                ex.Message + _NL + _NL +
                                "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void btn_SelectFileImport_DragDrop(object sender, DragEventArgs e)
        {
            string[] CaleFisiereSelectate = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (CaleFisiereSelectate.Count() > 1)
            {
                MessageBox.Show(this, "Wow, wow... too many files!" + _NL +
                                "I only need one...", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                btn_SelectFileImport_DragLeave(sender, e);
                return;
            }
            FileAttributes attr = File.GetAttributes(CaleFisiereSelectate[0]);
            if (attr.HasFlag(FileAttributes.Directory))
            {
                MessageBox.Show(this, "I needed a .zif file, and..." + _NL +
                                "I got a folder. :(", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                btn_SelectFileImport_DragLeave(sender, e);
                return;
            }
            string extensie = Path.GetExtension(CaleFisiereSelectate[0]);
            if (extensie != ".zif")
            {
                MessageBox.Show(this, $"Woops! This is a {extensie} file..." + _NL +
                                "I need a .zif file please!", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                btn_SelectFileImport_DragLeave(sender, e);
                return;
            }
            long fileLength = new FileInfo(CaleFisiereSelectate[0]).Length;
            if (fileLength == 0)
            {
                MessageBox.Show(this, $"It seems like this file is empty." + _NL +
                                "Please check it out, or try another one.", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                btn_SelectFileImport_DragLeave(sender, e);
                return;
            }
            btn_SelectFileImport_DragLeave(sender, e);
            ValideazaFisierImport(CaleFisiereSelectate[0]);
        }

        private void btn_SelectFileImport_DragEnter(object sender, DragEventArgs e)
        {
            btn_SelectFileImport.Text = "> Drop the file here <";
            btn_SelectFileImport.BackColor = Color.FromArgb(80, 80, 80);
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) { e.Effect = DragDropEffects.Copy; }
        }

        private void btn_SelectFileImport_DragLeave(object sender, EventArgs e)
        {
            btn_SelectFileImport.Text = "Select file (or drag && drop)";
            btn_SelectFileImport.BackColor = Color.FromArgb(46, 46, 46);
        }

        private void btn_SelectFileImport_Click(object sender, EventArgs e)
        {
            OpenFileD.Title = "Please select .zif a file to import";
            OpenFileD.Filter = "ZIF files (*.zif)|*.zif";
            OpenFileD.FilterIndex = 0;
            if (OpenFileD.ShowDialog(this) == DialogResult.OK)
            {
                long fileLength = new FileInfo(OpenFileD.FileName).Length;
                if (fileLength == 0)
                {
                    MessageBox.Show(this, $"It seems like this file is empty." + _NL +
                                    "Please check it out, or try another one.", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                ValideazaFisierImport(OpenFileD.FileName);
            }
        }

        private void ValideazaFisierImport(string caleFisier)
        {
            try
            {
                if (!File.Exists(caleFisier))
                {
                    MessageBox.Show(this, $"I can't find that file..." + _NL +
                                    "Maibe it was moved or deleted.", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                ReseteazaControaleImport();
                string dateFisier = File.ReadAllText(caleFisier);
                if (cls_Functii.ValideazaJSON(dateFisier))
                {
                    _exportImport = JsonConvert.DeserializeObject<ExportImport>(dateFisier);
                    GataDeImport();
                }
                else
                {
                    _dateFisierImport = dateFisier;
                    txt_PwDecrypt.Enabled = chk_ShowDecryptPw.Enabled = btn_DecryptData.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while validating this file:" + _NL +
                                ex.Message + _NL + _NL +
                                "Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chk_ShowDecryptPw_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_ShowDecryptPw.Checked)
            { txt_PwDecrypt.PasswordChar = '\0'; }
            else { txt_PwDecrypt.PasswordChar = '•'; }
        }

        private void btn_DecryptData_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txt_PwDecrypt.Text))
                {
                    MessageBox.Show(this, "This file is encrypted. Please enter the password.",
                                            _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (txt_PwDecrypt.TextLength < 4)
                {
                    MessageBox.Show(this, "Sorry, the password must be at least 4 characters.",
                                          _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (!DecripteazaFisier())
                {
                    MessageBox.Show(this, $"Wrong password!" + _NL +
                                    "Please try again.", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _exportImport = JsonConvert.DeserializeObject<ExportImport>(_dateFisierImport);
                MessageBox.Show(this, $"File decrypted successfully!" + _NL +
                                "Go to the next seps below.", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txt_PwDecrypt.Enabled = chk_ShowDecryptPw.Enabled = btn_DecryptData.Enabled = false;
                GataDeImport();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while decrypting the file:" + _NL +
                                ex.Message + _NL + _NL +
                                "Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_IMPORTdata_Click(object sender, EventArgs e)
        {
            try
            {
                if (!chk_ImportTemplates.Checked && !chk_ImportMeetings.Checked)
                {
                    MessageBox.Show(this, $"You have to choose at least one type of data:" + _NL +
                                    "Templates and/or meetings.", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (cmb_ImportMode.SelectedIndex == 1)
                {
                    if (chk_ImportTemplates.Checked) 
                    {
                        Setari.Preseturi.Clear();
                        Setari.Preseturi = _exportImport.Preseturi;
                    }
                    if (chk_ImportMeetings.Checked) 
                    {
                        Setari.Intruniri.Clear();
                        Setari.Intruniri = _exportImport.Intruniri;
                    }
                }
                else
                {
                    if (chk_ImportTemplates.Checked)
                    {
                        foreach (KeyValuePair<string, string> preset in _exportImport.Preseturi)
                        {
                            if (Setari.Preseturi.ContainsKey(preset.Key))
                            {
                                if (chk_ImportKeepDup.Checked)
                                {
                                    int copie = 1;
                                    string numeCopie = $"{preset.Key} (copy {copie})";
                                    while (Setari.Preseturi.ContainsKey(numeCopie))
                                    {
                                        copie++;
                                        numeCopie = $"{preset.Key} (copy {copie})";
                                    }
                                    Setari.Preseturi.Add(numeCopie, preset.Value);
                                }
                            }
                            else
                            {
                                Setari.Preseturi.Add(preset.Key, preset.Value);
                            }
                        }
                    }
                    //
                    if (chk_ImportMeetings.Checked)
                    {
                        foreach (Intrunire intrunire in _exportImport.Intruniri)
                        {
                            if (Setari.Intruniri.Contains(intrunire))
                            {
                                if (chk_ImportKeepDup.Checked)
                                {
                                    int copie = 1;
                                    Intrunire newIntrunire = new Intrunire()
                                    {
                                        Nume = $"{intrunire.Nume} (copy {copie})",
                                        Ziua = intrunire.Ziua,
                                        Ora = intrunire.Ora
                                    };
                                    while (Setari.Intruniri.Contains(newIntrunire))
                                    {
                                        copie++;
                                        newIntrunire.Nume = $"{intrunire.Nume} (copy {copie})";
                                    }
                                    Setari.Intruniri.Add(newIntrunire);
                                }
                            }
                            else
                            {
                                Setari.Intruniri.Add(intrunire);
                            }
                        }
                    }
                }
                RaspunsFunctie rf = cls_Setari.Salveaza(cls_Variabile.CaleDateAPP + "\\APP.set");
                if (rf.Eroare)
                {
                    MessageBox.Show(this, "Error saving the settings:" + _NL +
                                    rf.Mesaj + _NL + _NL +
                                    "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
                else
                {
                    btn_ResetIntruniri_Click(null, null);
                    ReseteazaControaleImport();
                    cls_Functii.IncarcaPreseturi(cmb_Preseturi, lbl_Presets);
                    MessageBox.Show(this, "Data imported and saved successfully! :)",
                                           _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error while importing the data:" + _NL +
                                ex.Message + _NL + _NL +
                                "The app will close. Please contact AvA.Soft", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void ReseteazaControaleImport()
        {
            btn_SelectFileImport.Enabled = true;
            txt_PwDecrypt.Enabled = chk_ShowDecryptPw.Enabled = chk_ShowDecryptPw.Checked = btn_DecryptData.Enabled = chk_ImportTemplates.Enabled = chk_ImportTemplates.Checked = chk_ImportMeetings.Enabled = chk_ImportMeetings.Checked = cmb_ImportMode.Enabled = chk_ImportKeepDup.Enabled = chk_ImportKeepDup.Checked = btn_IMPORTdata.Enabled = false;
            cmb_ImportMode.SelectedIndex = 0;
            chk_ImportTemplates.Text = $"Import templates";
            chk_ImportMeetings.Text = $"Import meetings";
            txt_PwDecrypt.Text = "";
        }

        private void GataDeImport()
        {
            if (_exportImport.Preseturi.Count == 0 & _exportImport.Intruniri.Count == 0)
            {
                MessageBox.Show(this, $"There are no templates and no meetings in this file." + _NL +
                                "Please try another one. Sorry!", _TitluAPP, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ReseteazaControaleImport();
                return;
            }
            chk_ImportTemplates.Text = $"Import templates: {_exportImport.Preseturi.Count}";
            chk_ImportMeetings.Text = $"Import meetings: {_exportImport.Intruniri.Count}";
            if (_exportImport.Preseturi.Count > 0) { chk_ImportTemplates.Enabled = chk_ImportTemplates.Checked = true; }
            else { chk_ImportTemplates.Checked = false; }
            if (_exportImport.Intruniri.Count > 0) { chk_ImportMeetings.Enabled = chk_ImportMeetings.Checked = true; }
            else { chk_ImportMeetings.Checked = false; }
            cmb_ImportMode.Enabled = chk_ImportKeepDup.Enabled = btn_IMPORTdata.Enabled = true;
        }

        private bool DecripteazaFisier()
        {
            try
            {
                _dateFisierImport = cls_Cryptography.Decrypt(_dateFisierImport, txt_PwDecrypt.Text);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void lbl_Mexport_Click(object sender, EventArgs e)
        {
            pnl_Export.Visible = true;
            pnl_Import.Visible = false;
            lbl_Mexport.Font = _lblActiv;
            lbl_Mimport.Font = _lblInactiv;
            lbl_Mexport.ForeColor = Color.White;
            lbl_Mimport.ForeColor = Color.Gray;
        }

        private void lbl_Mimport_Click(object sender, EventArgs e)
        {
            pnl_Export.Visible = false;
            pnl_Import.Visible = true;
            lbl_Mexport.Font = _lblInactiv;
            lbl_Mimport.Font = _lblActiv;
            lbl_Mexport.ForeColor = Color.Gray;
            lbl_Mimport.ForeColor = Color.White;
        }

        private void btn_ShowExportImport_Click(object sender, EventArgs e)
        {
            btn_ShowExportImport.Visible = false;
            pnl_ExportImportData.Visible = true;
        }

        private void btn_SET_MouseEnter(object sender, EventArgs e)
        {
            if (pnl_SET.Visible)
            {
                lbl_Runs.ForeColor = Color.FromArgb(255, 126, 126);
            }
        }

        private void btn_SET_MouseLeave(object sender, EventArgs e)
        {
            lbl_Runs.ForeColor = Color.Gray;
        }

    }
}

#region KEYBOARD HOOK
#region V1
//_hookID = SetHook(_proc);
//UnhookWindowsHookEx(_hookID);

//private const int WH_KEYBOARD_LL = 13;
//private const int WM_KEYDOWN = 0x0100;
//private const int WM_KEYUP = 0x0101;
//private static LowLevelKeyboardProc _proc = HookCallback;
//private static IntPtr _hookID = IntPtr.Zero;

//private static IntPtr SetHook(LowLevelKeyboardProc proc)
//{
//    using (Process curProcess = Process.GetCurrentProcess())
//    using (ProcessModule curModule = curProcess.MainModule)
//    {
//        return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
//    }
//}

//private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

//private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
//{
//    int vkCode = Marshal.ReadInt32(lParam);
//    if (nCode >= 0 && (vkCode == 162 || vkCode == 163)) // CTRL
//    {
//        if (wParam == (IntPtr)WM_KEYDOWN)
//        {
//            _isCtrlDown = true;

//        }
//        else if (wParam == (IntPtr)WM_KEYUP)
//        {
//            _isCtrlDown = false;
//        }
//    }
//    return CallNextHookEx(_hookID, nCode, wParam, lParam);
//    //if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
//    //{
//    //    int vkCode = Marshal.ReadInt32(lParam);
//    //    if (vkCode == 162 || vkCode == 163)
//    //    {
//    //    }
//    //    //Console.WriteLine("Down: " + (Keys)vkCode);
//    //}
//    //if (nCode >= 0 && wParam == (IntPtr)WM_KEYUP)
//    //{
//    //    int vkCode = Marshal.ReadInt32(lParam);
//    //    if (vkCode == 162 || vkCode == 163)
//    //    {
//    //    }
//    //    //Console.WriteLine("UP: " + (Keys)vkCode);
//    //}
//}

//[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

//[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//[return: MarshalAs(UnmanagedType.Bool)]
//private static extern bool UnhookWindowsHookEx(IntPtr hhk);

//[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

//[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
//private static extern IntPtr GetModuleHandle(string lpModuleName);
#endregion
#endregion