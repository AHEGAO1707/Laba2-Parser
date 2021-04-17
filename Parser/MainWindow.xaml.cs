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
    public partial class MainWindow : Window
    {

        public static int page = 0;

        public MainWindow()
        {
            InitializeComponent();

            if (System.IO.File.Exists("thrlist.xlsx"))
            {
                MessageBox.Show("Файл локальной базы данных угроз уже существует. Дальнейшие данные будут браться из него.", "Файл уже существует!", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("Файл локальной базы данных угроз не существует. Вы хотите скачать его?", "Скачать файл?", MessageBoxButton.YesNo, MessageBoxImage.Question);
                switch (result)
                {
                    case MessageBoxResult.No:
                        MessageBox.Show("Не очень то уж и хотелось.. До свидания!", "Ну и ладно..", MessageBoxButton.OK, MessageBoxImage.Information);
                        Environment.Exit(0);
                        break;
                    case MessageBoxResult.Yes:
                        MessageBox.Show("Супер! Скачиваю файл.", "Скачивание...", MessageBoxButton.OK, MessageBoxImage.Information);
                        WorkWithFiles.Zagruzka();
                        break;
                }
            }

            ZaplInfo(this.listView, ButBack, ButNext, lbl);
        }

        public static void ZaplInfo(ListView lv, Button b, Button b1, Label lb) //заполнение первой страницы
        {
            page = 0;
            lb.Content = page + 1;
            var gridView = new GridView();
            lv.View = gridView;
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Идентификатор угрозы",
                Width = 150,
                DisplayMemberBinding = new Binding("Id")
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Наименование угрозы",
                Width = 1000,
                DisplayMemberBinding = new Binding("Name")
            });

            WorkWithFiles.ExcelToList(false);
            lv.Items.Clear();
            for (int i = 0; i < 15; i++)
            {
                lv.Items.Add(new ItemInList { Id = "УБИ." + InfoParser.list1[i][1], Name = InfoParser.list1[i][2] });
            }
            b.Visibility = Visibility.Hidden;
            b1.Visibility = Visibility.Visible;
        }

        public class ItemInList // чтобы получилось удобно засунуть данные в листвью
        {
            public string Id { get; set; }
            public string Name { get; set; }
        }

        private void ButNext_Click(object sender, RoutedEventArgs e) //кнопка вперёд
        {
            page++;
            lbl.Content = page + 1;
            if (page != 0)
            {
                ButBack.Visibility = Visibility.Visible;
            }
            this.listView.Items.Clear();
            for (int i = 0; i < 15; i++)
            {
                if (InfoParser.list1.Count > i + page * 15)
                {
                    this.listView.Items.Add(new ItemInList { Id = "УБИ." + InfoParser.list1[i + page * 15][1], Name = InfoParser.list1[i + page * 15][2] });
                }
                else
                {
                    ButNext.Visibility = Visibility.Hidden;
                }
            }
        }

        private void ButBack_Click(object sender, RoutedEventArgs e) //кнопка назад
        {
            page--;
            lbl.Content = page + 1;
            if (page != 0)
            {
                ButNext.Visibility = Visibility.Visible;
            }
            else
            {
                ButBack.Visibility = Visibility.Hidden;
            }
            this.listView.Items.Clear();
            for (int i = 0; i < 15; i++)
            {
                if (InfoParser.list1.Count > i + page * 15)
                {
                    this.listView.Items.Add(new ItemInList { Id = "УБИ." + InfoParser.list1[i + page * 15][1], Name = InfoParser.list1[i + page * 15][2] });
                }
            }
        }

        public static string ZamenaBinarn(string s) //потому что Replace работает только с чарами, а "нет" и "да" туда не всунуть
        {
            if (s == "0")
            {
                return "нет";
            }
            else
            {
                return "да";
            }
        }

        private void listView_SelectionChanged(object sender, SelectionChangedEventArgs e) //отображение информации о выбраной записи
        {
            if (listView.SelectedItem is ItemInList x)
            {
                MessageBox.Show("Идентификатор угрозы: " + InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][1] +
                    "\r\n Наименование угрозы: " + InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][2] +
                    "\r\n Описание угрозы: " + InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][3] +
                    "\r\n Источник угрозы: " + InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][4] +
                    "\r\n Объект воздействия угрозы: " + InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][5] +
                    "\r\n Нарушение конфиденциальности: " + ZamenaBinarn(InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][6]) +
                    "\r\n Нарушение целостности: " + ZamenaBinarn(InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][7]) +
                    "\r\n Нарушение доступности: " + ZamenaBinarn(InfoParser.list1[int.Parse(x.Id.Remove(0, 4)) - 1][8]), "Все сведения о угрозе " + x.Id, MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ButSave_Click(object sender, RoutedEventArgs e) //кнопка сохранить
        {
            MessageBoxResult askAboutSave = MessageBox.Show("Сохранить все сведения локальной базы данных угроз безопасности в .txt файле?", "Сохранить?", MessageBoxButton.YesNo, MessageBoxImage.Question);
            switch (askAboutSave)
            {
                case MessageBoxResult.No:
                    MessageBox.Show("Зачем тогда на кнопку надо было нажимать..", "На нет и суда нет!", MessageBoxButton.OK, MessageBoxImage.Information);
                    break;
                case MessageBoxResult.Yes:
                    WorkWithFiles.TxtCreate();
                    WorkWithFiles.FromListToTxt(InfoParser.list1);
                    MessageBox.Show("Супер! Был создан .txt файл и вся информация локальной базы данных была записана в него!", "Успешно!", MessageBoxButton.OK, MessageBoxImage.Information);
                    break;
            }
        }

        private void ButUpdt_Click(object sender, RoutedEventArgs e) //кнопка обновить данные
        {
            MessageBoxResult askAboutSave = MessageBox.Show("Обновить сведения локальной базы данных угроз безопасности информации?", "Обновить?", MessageBoxButton.YesNo, MessageBoxImage.Question);
            switch (askAboutSave)
            {
                case MessageBoxResult.No:
                    MessageBox.Show("Зачем тогда на кнопку надо было нажимать..", "На нет и суда нет!", MessageBoxButton.OK, MessageBoxImage.Information);
                    break;
                case MessageBoxResult.Yes:
                    MessageBoxResult askAboutWhat = MessageBox.Show("Вы хотите произвести сравнение с тем файлом, что уже скачан? (Вы, возможно, его изменили)." +
                        "Или хотите сравнить с файлом, который будет скачан из банка данных угроз ФСТЭК России?" + "\r\n" +
                        "Нажмите да, если \"первое\" или нет, если \"второе\", пожалуйста :)", "What do u want?", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                    switch (askAboutWhat)
                    {
                        case MessageBoxResult.No:
                            WorkWithFiles.Zagruzka();
                            break;
                        case MessageBoxResult.Yes:
                            break;
                        case MessageBoxResult.Cancel:
                            MessageBox.Show("А из Вас хороший тестировщик", "Отмена!", MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                    }

                    MessageBox.Show("Пытаюсь обновить...", "Попытка не пытка!", MessageBoxButton.OK, MessageBoxImage.Information);

                    try
                    {
                        WorkWithFiles.ExcelToList(true);
                        var d = InfoParser.InfPars(InfoParser.list1, InfoParser.list2);
                        if (d.Count == 0)
                        {
                            MessageBox.Show("Ничего не изменилось!", "Zero chages!", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                        {
                            MessageBox.Show("Всего записей (строк) обновлено " + d.Count, "Count!", MessageBoxButton.OK, MessageBoxImage.Information);
                            foreach (var item in d)
                            {
                                MessageBox.Show(item.Value, "Идентификатор угрозы: " + item.Key, MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    break;
            }
            InfoParser.list1.Clear();
            InfoParser.list1.AddRange(InfoParser.list2);
            InfoParser.list2.Clear();
            ZaplInfo(this.listView, this.ButBack, this.ButNext, this.lbl); //тут надо первую страничку снова открыть и заполнить новыми данными
        }
    }
}

