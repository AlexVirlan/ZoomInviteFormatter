using Newtonsoft.Json.Schema;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace ZoomInviteFormatter.Clase
{
    public class cls_Functii
    {
        public static void ReseteazaPreseturi()
        {
            Setari.Preseturi = GetDataFromResourceFile("PreseturiDefault");
        }

        public static Dictionary<string, string> GetDataFromResourceFile(string resFileName = "PreseturiDefault")
        {
            if (string.IsNullOrEmpty(resFileName)) { return new Dictionary<string, string>(); }
            try
            {
                string dateChangeLog = "";
                Dictionary<string, string> result = new Dictionary<string, string>();
                Assembly assembly = Assembly.GetExecutingAssembly();
                IEnumerable<string> resources = assembly.GetManifestResourceNames().Where(x => x.Contains(resFileName)).OrderBy(o => o);
                foreach (string res in resources)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(res))
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        dateChangeLog = reader.ReadToEnd();
                    }
                    result.Add(res.Replace(".txt", "").Replace("ZoomInviteFormatter.PreseturiDefault.", ""), dateChangeLog);
                }
                return result;
            }
            catch (Exception ex)
            {
                return new Dictionary<string, string>();
            }
        }

        public static void IncarcaPreseturi(ComboBox comboBox, Label lbl, int selectIndex = 0)
        {
            comboBox.Items.Clear();
            foreach (KeyValuePair<string, string> P in Setari.Preseturi)
            {
                comboBox.Items.Add(P.Key);
            }
            if ((comboBox.Items.Count - 1) >= selectIndex) { comboBox.SelectedIndex = selectIndex; }
            lbl.Text = $"Templates ({comboBox.Items.Count}):";
        }

        public static ComboboxItem GetDayFromDayNumber(int dayNumber)
        {
            switch (dayNumber)
            {
                case 1: return new ComboboxItem() { Text = "1. Monday", Value = "1" };
                case 2: return new ComboboxItem() { Text = "2. Tuesday", Value = "2" };
                case 3: return new ComboboxItem() { Text = "3. Wednesday", Value = "3" };
                case 4: return new ComboboxItem() { Text = "4. Thursday", Value = "4" };
                case 5: return new ComboboxItem() { Text = "5. Friday", Value = "5" };
                case 6: return new ComboboxItem() { Text = "6. Saturday", Value = "6" };
                case 7: return new ComboboxItem() { Text = "7. Sunday", Value = "7" };
                default: return null;
            }
        }

        public static int GetDayNumberFromDayName(string dayName)
        {
            switch (dayName)
            {
                case "Monday": return 1;
                case "Tuesday": return 2;
                case "Wednesday": return 3;
                case "Thursday": return 4;
                case "Friday": return 5;
                case "Saturday": return 6;
                case "Sunday": return 7;
                default: return 0;
            }
        }

        public static void FillComboWithWeekDays(ComboBox comboBox)
        {
            comboBox.Items.Add(GetDayFromDayNumber(1));
            comboBox.Items.Add(GetDayFromDayNumber(2));
            comboBox.Items.Add(GetDayFromDayNumber(3));
            comboBox.Items.Add(GetDayFromDayNumber(4));
            comboBox.Items.Add(GetDayFromDayNumber(5));
            comboBox.Items.Add(GetDayFromDayNumber(6));
            comboBox.Items.Add(GetDayFromDayNumber(7));
        }

        public static bool ValideazaJSON(string jsonData)
        {
            try
            {
                JSchema schemaJson = JSchema.Parse(jsonData);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

    }

    public static class Extensii
    {
        public static void RenameKey<TKey, TValue>(this IDictionary<TKey, TValue> dic, TKey fromKey, TKey toKey)
        {
            TValue value = dic[fromKey];
            dic.Remove(fromKey);
            dic[toKey] = value;
        }

        public static void HighlightText(this RichTextBox myRtb, string word, Color color)
        {
            if (string.IsNullOrEmpty(word)) { return; }
            int s_start = myRtb.SelectionStart, startIndex = 0, index;
            while ((index = myRtb.Text.IndexOf(word, startIndex)) != -1)
            {
                myRtb.Select(index, word.Length);
                myRtb.SelectionColor = color;
                startIndex = index + word.Length;
            }
            myRtb.SelectionStart = s_start;
            myRtb.SelectionLength = 0;
            myRtb.SelectionColor = myRtb.ForeColor;
        }
    }
}
