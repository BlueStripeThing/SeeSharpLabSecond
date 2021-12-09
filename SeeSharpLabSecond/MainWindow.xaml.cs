using System;
using System.Collections.Generic;
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
using System.IO;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeeSharpLabSecond
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        public  List<Threat> dbase  = new List<Threat>();
        //private List<BetterThreat> betterDB = new List<BetterThreat>();
        private string path;
        private int pageCounter = 1;
        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            ImportInfo();
            dbase = GetBase(path);
            Pagination(1);
        }

        //Поиск существует ли файл в указанной папке. Если нет - скачать файл и положить в указанную папку.
        private void ImportInfo()
        {
            DirectoryInfo dirInfo = new DirectoryInfo(PathTBox.Text);
            if (!dirInfo.Exists)
            {
                MessageBox.Show("Такой папки не существует", "Ошибка");
            }
            else
            {
                path = PathTBox.Text + @"\thrlist.xlsx";
                string[] thrlist = Directory.GetFiles(PathTBox.Text, "thrlist.xlsx");
                ExcelConnection excel;
                bool created = false;
                if (thrlist.Length != 0)
                {
                    foreach (var a in thrlist)
                    {
                        if (a == "thrlist.xlsx") excel = new ExcelConnection(a);
                        created = true;
                    }
                }
                if (!created)
                {
                    WebClient webClient = new WebClient();
                    string link = @"https://bdu.fstec.ru/files/documents/thrlist.xlsx";
                    webClient.DownloadFile(new Uri(link), path);
                }
                IdColumn.Visibility = Visibility.Visible;
                NameColumn.Visibility = Visibility.Visible;
                NextBtn.Visibility = Visibility.Visible;
                PrevBtn.Visibility = Visibility.Visible;
                PageBlock.Visibility = Visibility.Visible;
                PageBlock.Text ="1 из " + (dbase.Count / 15 + 1);
                pageCounter = 1;
            }
        }

        //Получение списка угроз из файла
        public static List<Threat> GetBase(string path)
        {
            List<Threat> tempBase = new List<Threat>();
            try
            {
                ExcelConnection excel = new ExcelConnection(path);
                
                var countQuery = from a in excel.UrlConnexion.WorksheetNoHeader("Sheet") select a;
                var query = from a in excel.UrlConnexion.WorksheetRange<Threat>("A2", "J" + (countQuery.Count()).ToString(), "Sheet") select a;

                foreach (var row in query)
                {
                    tempBase.Add(row);
                }

            }
            catch (System.Data.OleDb.OleDbException exc)
            {
                MessageBox.Show($"Произошла ошибка. Сообщение об ошибке = {exc.Message}");
            }
            return tempBase;
        }

        //Обновление базы данных из файла
        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(PathTBox.Text);
            if (!dirInfo.Exists)
            {
                MessageBox.Show("Такой папки не существует", "Ошибка");
            }
            else
            {
                path = PathTBox.Text + @"\thrlist.xlsx";
                ComparisonWind comparison = new ComparisonWind();
                bool flag = true;
                try
                {
                    ImportInfo();
                    comparison.UpdateWindows(dbase, GetBase(path));
                    dbase = GetBase(path);
                }
                catch (IOException exc)
                {
                    flag = false;
                    MessageBox.Show($"Произошла ошибка {path}. Сообщение об ошибке = {exc.Message}");
                }
                if (flag)
                {
                    
                    int i = comparison.CompareBases();
                    if (i != 0)
                    {
                        MessageBoxResult result = MessageBox.Show($"Обновление прошло успешно. Обновлено {i}. Посмотреть изменения?", "Обновление", MessageBoxButton.YesNo);
                        if (result == MessageBoxResult.Yes)
                        {
                            comparison.Show();
                        }
                        
                    }
                    else MessageBox.Show("Изменений нет");
                }
                Pagination(1);
                pageCounter = 1;
            }
        }
        //Получение данных о записи по нажатию
        private void BetterGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BetterThreat betterThreat = BetterGrid.SelectedItem as BetterThreat;
            Threat threat = new Threat();
            foreach (Threat th in dbase)
            {
                if (th.Id == betterThreat.GetId()) threat = th;
            }
            IdBlock.Text = "ID: " + threat.Id.ToString();
            NameBlock.Text = "Наименование: " + threat.Name;
            DescriptionBlock.Text = "Описание: " + threat.Description;
            TargetBlock.Text = "Источник угрозы: " + threat.Target;
            SourceBlock.Text = "Объект воздействи: " + threat.Source;
            ConfBlock.Text = "Нарушение конфиденциальности: ";
            if (threat.Confidence) ConfBlock.Text += "Да";
            else ConfBlock.Text += "Нет";
            IntegrBlock.Text = "Нарушение целостности: ";
            if (threat.Integrity) IntegrBlock.Text += "Да";
            else IntegrBlock.Text += "Нет";
            AvailBlock.Text = "Нарушение доступности: ";
            if (threat.Availability) AvailBlock.Text += "Да";
            else AvailBlock.Text += "Нет";
            AddedBlock.Text = "Дата добавления: " + threat.AddedDate.ToString("dd/MM/yyyy");
            ChangedBlock.Text = "Дата изменения: " + threat.ChangedDate.ToString("dd/MM/yyyy");

        }
        //Сохранение в файле
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show($"Да - сохранить в формате xlsx. Нет - в .txt", "Сохранение", MessageBoxButton.YesNo);
            if (result == MessageBoxResult.Yes)
            {
                ExcelSave();
            }
            if (result == MessageBoxResult.No)
            {
                path = PathTBox.Text + @"\База.txt";
                using (FileStream fstream = new FileStream(path, FileMode.OpenOrCreate))
                {

                    foreach (Threat th in dbase)
                    {
                        byte[] array = System.Text.Encoding.Default.GetBytes(th.ToString());

                        fstream.Write(array, 0, array.Length);

                    }
                }
            }
        }
        //Сохранение в excel
        private void ExcelSave()
        {
            path = PathTBox.Text + @"\База в excel";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
            excelWorkSheet.Name = "Тест";
            excelWorkSheet.Cells[1, 1] = "Идентификатор УБИ";
            excelWorkSheet.Cells[1, 2] = "Наименование УБИ";
            excelWorkSheet.Cells[1, 3] = "Описание";
            excelWorkSheet.Cells[1, 4] = "Источник угрозы (характеристика и потенциал нарушителя)";
            excelWorkSheet.Cells[1, 5] = "Объект воздействия";
            excelWorkSheet.Cells[1, 6] = "Нарушение конфиденциальности";
            excelWorkSheet.Cells[1, 7] = "Нарушение целостности";
            excelWorkSheet.Cells[1, 8] = "Нарушение доступности";
            excelWorkSheet.Cells[1, 9] = "Дата включения угрозы в БнД УБИ";
            excelWorkSheet.Cells[1, 10] = "Дата последнего изменения данных";
            int counter = 2;
            excelWorkSheet.Columns["A:J"].ColumnWidth = 15.00;
            excelWorkSheet.Columns["B:E"].ColumnWidth = 33.00;
            excelWorkSheet.Rows.RowHeight = 20.00;
            foreach (Threat th in dbase)
            {
                excelWorkSheet.Cells[counter,1] = th.Id;
                excelWorkSheet.Cells[counter, 2] = th.Name;
                excelWorkSheet.Cells[counter, 3] = th.Description;
                excelWorkSheet.Cells[counter, 4] = th.Source;
                excelWorkSheet.Cells[counter, 5] = th.Target;
                if (th.Confidence) excelWorkSheet.Cells[counter, 6] = "1";
                else excelWorkSheet.Cells[counter, 6] = "0";
                IntegrBlock.Text = "Нарушение целостности: ";
                if (th.Integrity) excelWorkSheet.Cells[counter,7] = "1";
                else excelWorkSheet.Cells[counter, 7] = "0";
                AvailBlock.Text = "Нарушение доступности: ";
                if (th.Availability) excelWorkSheet.Cells[counter, 8] = "1";
                else excelWorkSheet.Cells[counter, 8] = "0";
                excelWorkSheet.Cells[counter, 9] = th.AddedDate.ToString("dd/MM/yyyy");
                excelWorkSheet.Cells[counter, 10] = th.ChangedDate.ToString("dd/MM/yyyy");


                counter++;
            }
            excelWorkBook.SaveAs(path);
            excelWorkBook.Close();
            excelApp.Quit();
        }

        private void Pagination(int page)
        {
            PageBlock.Text = page + " из " + (dbase.Count/15+1);
            if (page == 1) PrevBtn.IsEnabled = false;
            else PrevBtn.IsEnabled = true;
            if (page == (dbase.Count / 15 + 1)) NextBtn.IsEnabled = false;
            else NextBtn.IsEnabled = true;
            List<BetterThreat> update = new List<BetterThreat>();
            if ((dbase.Count+1) / 15 == 0)
            {
                foreach (Threat th in dbase)
                {
                    update.Add(new BetterThreat("УБИ." + th.Id, th.Name));
                }
            }
            else
            {
                if (page * 15 < dbase.Count)
                {
                    for (int i = (page - 1) * 15; i < page * 15; i++)
                    {
                        update.Add(new BetterThreat("УБИ." + dbase[i].Id, dbase[i].Name));
                    }
                }
                else
                {
                    for (int i = (page - 1) * 15; i < dbase.Count; i++)
                    {
                        update.Add(new BetterThreat("УБИ." + dbase[i].Id, dbase[i].Name));
                    }
                }
            }
            BetterGrid.ItemsSource = update;
        }

        private void NextBtn_Click(object sender, RoutedEventArgs e)
        {
            Pagination(++pageCounter);
        }

        private void PrevBtn_Click(object sender, RoutedEventArgs e)
        {
            Pagination(--pageCounter);
        }
    }
}


