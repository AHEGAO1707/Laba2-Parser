using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
namespace Parser
{
    public class InfoParser
    {
        static string saveWhere = Directory.GetCurrentDirectory(); // где сохранен.
        public static List<List<string>> list1 = new List<List<string>>();
        public static List<List<string>> list2 = new List<List<string>>();

        public static void FirstList(string s)  //добавляем в первый лист
        {
            list1.Add(s.Split('|').ToList<string>());
        }

        public static void SecondList(string s) //добавляем вo второй лист
        {
            list2.Add(s.Split('|').ToList<string>());
        }

        public static Dictionary<string, string> InfPars(List<List<string>> l1, List<List<string>> l2) //возвращает словарь который содержит все изменненные записи
        {
            Dictionary<string, string> DictWhatChanged = new Dictionary<string, string>();
            for (int i = 0; i < l1.Count; i++)
            {
                for (int j = 0; j < l1[i].Count; j++)
                {
                    if (l1[i][j] != l2[i][j])
                    {
                        if (DictWhatChanged.ContainsKey(l2[i][1]))
                        {
                            string key = l2[i][1];
                            DictWhatChanged[key] += " БЫЛО: " + l1[i][j] + "\r\n" + " СТАЛО: " + l2[i][j] + "\r\n";
                        }
                        else
                        {
                            DictWhatChanged.Add(l2[i][1], "БЫЛО: " + l1[i][j] + "\r\n" + " СТАЛО: " + l2[i][j] + "\r\n");
                        }
                    }
                }
            }
            return DictWhatChanged;
        }
    }
}
