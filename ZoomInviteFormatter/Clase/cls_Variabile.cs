using System;
using System.Collections.Generic;

namespace ZoomInviteFormatter.Clase
{
    public class cls_Variabile
    {
        public static string CaleDateAPP = "";
        public static bool PrimaRulare = false;
        public static string VersiuneApp = "1.0";
        public static string ContactMail = "dj.3agl3@gmail.com";
        public static string ContactTelegram = "AvA_Alex";
        public static string ContactWebsite = "";
        public static string FisierWatcher = ".watcher";
        public static bool StartedByWatcher = false;
    }

    public class Intrunire : IEquatable<Intrunire>
    {
        public string Nume { get; set; }
        public int Ziua { get; set; }
        public string Ora { get; set; }

        public bool Equals(Intrunire other)
        {
            return this.Nume == other.Nume &&
                   this.Ziua == other.Ziua &&
                   this.Ora == other.Ora;
        }
    }
    
    public class ExportImport
    {
        public Dictionary<string, string> Preseturi = new Dictionary<string, string>();
        public List<Intrunire> Intruniri = new List<Intrunire>();
    }

    public class ComboboxItem
    {
        public string Text { get; set; }
        public string Value { get; set; }
        public override string ToString()
        {
            return Text;
        }
    }

    public class RaspunsFunctie
    {
        public bool Eroare { get; set; }
        public string Mesaj { get; set; }
        public string StackTrace { get; set; }
    }
}
