using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace ZoomInviteFormatter.Clase
{
    class Setari
    {
        #region Aplicatie
        public static Dictionary<string, string> Preseturi = new Dictionary<string, string>();
        public static List<Intrunire> Intruniri = new List<Intrunire>();
        public static int AppRuns = 0;
        public static string TimeMode = "next"; // next, nearest
        public static bool RememberLastTemplate = true;
        public static int LastTemplateIndex = 0;
        public static double Trans = 1;
        public static bool KeepOnTop = false;
        #endregion
    }

    [Serializable]
    class cls_Setari : Setari
    {
        public Dictionary<string, string> _Preseturi { get { return Preseturi; } set { Preseturi = value; } }
        public List<Intrunire> _Intruniri { get { return Intruniri; } set { Intruniri = value; } }
        public int _AppRuns { get { return AppRuns; } set { AppRuns = value; } }
        public string _TimeMode { get { return TimeMode; } set { TimeMode = value; } }
        public bool _RememberLastTemplate { get { return RememberLastTemplate; } set { RememberLastTemplate = value; } }
        public int _LastTemplateIndex { get { return LastTemplateIndex; } set { LastTemplateIndex = value; } }
        public double _Trans { get { return Trans; } set { Trans = value; } }
        public bool _KeepOnTop { get { return KeepOnTop; } set { KeepOnTop = value; } }

        public static RaspunsFunctie Salveaza(string caleFisier = "APP.set")
        {
            string dateSetari = "", cryptoPass = "2.A.v.A.6";
            //cryptoPass = "AAv - " + cls_AmprentaPC.Amprenta(folosesteCache: true, hashLocal: false, delimitatorH: "") + " - 17";
            //cryptoPass = "2_" + cls_Hash.MD5_Hash(cryptoPass) + "_6";
            dateSetari = JsonConvert.SerializeObject(new cls_Setari());
            dateSetari = cls_Cryptography.Encrypt(dateSetari, cryptoPass);
            try
            {
                File.WriteAllText(caleFisier, dateSetari);
            }
            catch (Exception ex)
            {
                //Thread tErrLog = new Thread(() => cls_Functii.LOG(cls_Functii.TipLOG.ERROR, ex.Message));
                //tErrLog.Start();
                return new RaspunsFunctie() { Eroare = true, Mesaj = ex.Message, StackTrace = ex.StackTrace };
            }
            return new RaspunsFunctie() { Eroare = false };
        }

        public static RaspunsFunctie Incarca(string caleFisier = "APP.set")
        {
            if (!File.Exists(caleFisier))
                return new RaspunsFunctie() { Eroare = true, Mesaj = "The settings file is missing from:" + Environment.NewLine + caleFisier };
            string dateSetari = "", cryptoPass = "2.A.v.A.6";
            //cryptoPass = "AAv - " + cls_AmprentaPC.Amprenta(folosesteCache: false, hashLocal: false, delimitatorH: "") + " - 17";
            //cryptoPass = "2_" + cls_Hash.MD5_Hash(cryptoPass) + "_6";
            try
            {
                dateSetari = File.ReadAllText(caleFisier);
                dateSetari = cls_Cryptography.Decrypt(dateSetari, cryptoPass);
                cls_Setari clasa_S = JsonConvert.DeserializeObject<cls_Setari>(dateSetari);
                //CopieDateClasa(new cls_Setari(), clasa_S);
            }
            catch (Exception ex)
            {
                //cls_Functii.LOG(cls_Functii.TipLOG.ERROR, ex.Message + Environment.NewLine + ex.StackTrace);
                return new RaspunsFunctie() { Eroare = true, Mesaj = ex.Message, StackTrace = ex.StackTrace };
            }
            return new RaspunsFunctie() { Eroare = false };
        }
    }
}
