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
    public class WorkWithFiles
    {
        static string link = "https://bdu.fstec.ru/files/documents/thrlist.xlsx"; // ссылка на файл
        static string saveWhere = Directory.GetCurrentDirectory(); // куда сохранять.

        public static void Zagruzka() //метод загрузки файла
        {
            using (var client = new System.Net.WebClient())
            {
                client.DownloadFile(new Uri(link), System.IO.Path.Combine(saveWhere, "thrlist.xlsx"));
            }
        }

        public static void ExcelToList(bool b) //добавление данных в лист
        {
            if (!b)
            {
                InfoParser.list1.Clear();
            }
            String filePath = System.IO.Path.Combine(saveWhere, "thrlist.xlsx");

            IWorkbook workbook = null;
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            workbook = new XSSFWorkbook(fs);

            ISheet sheet = workbook.GetSheetAt(0);

            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;
                string s = "";
                for (int i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        IRow currentRow = sheet.GetRow(i);

                        var cellVal0 = currentRow.GetCell(0).NumericCellValue;
                        var cellVal1 = currentRow.GetCell(1).StringCellValue;
                        var cellVal2 = currentRow.GetCell(2).StringCellValue;
                        var cellVal3 = currentRow.GetCell(3).StringCellValue;
                        var cellVal4 = currentRow.GetCell(4).StringCellValue;
                        var cellVal5 = currentRow.GetCell(5).NumericCellValue;
                        var cellVal6 = currentRow.GetCell(6).NumericCellValue;
                        var cellVal7 = currentRow.GetCell(7).NumericCellValue;

                        s = "|" + cellVal0 + "|" + cellVal1 + "|" + cellVal2 + "|" + cellVal3 + "|" + cellVal4 + "|" + cellVal5 + "|" + cellVal6 + "|" + cellVal7;

                        if (!b)
                        {
                            InfoParser.FirstList(s);
                        }
                        else
                        {
                            InfoParser.SecondList(s);
                        }
                    }
                    catch (Exception exx)
                    {
                        MessageBox.Show(exx.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error); //в основном ошибки о неправильном типе cell (не может выгрузить)
                    }
                }
            }
        }
        public static void TxtCreate() //создание txt файла
        {
            var TxtFile = System.IO.File.Create(System.IO.Path.Combine(saveWhere, "thrlist.txt"));
            TxtFile.Close();
        }

        public static void FromListToTxt(List<List<string>> list) //из листа в тхт
        {
            for (int i = 0; i < list.Count; i++)
            {
                foreach (string item in list[i])
                {
                    AddRowToTxt(item);
                }
            }
        }

        public static void AddRowToTxt(string s) //добавление строки в txt файл
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.Path.Combine(saveWhere, "thrlist.txt"), true))
            {
                file.WriteLine(s);
                file.Close();
            }
        }
    }
}
