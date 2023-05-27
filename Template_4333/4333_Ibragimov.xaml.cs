using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Text.Json;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text;
using System.Security.Cryptography;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Ibragimov.xaml
    /// </summary>
    public partial class _4333_Ibragimov : Window
    {
        public _4333_Ibragimov()
        {
            InitializeComponent();
        }

        private string GetHashString(string s)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(s);
            MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            string hash = "";
            foreach (byte b in byteHash)
            {
                hash += string.Format("{0:x2}", b);
            }
            return hash;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Список.xlsx)|*.xlsx",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new
            Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (isrpoEntities db = new isrpoEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    db.Workers.Add(new Workers()
                    {
                        Id= i,
                        RoleName = list[i, 0],
                        FIO = list[i, 1],
                        LoginName = list[i, 2],
                        PasswordName = list[i, 3],
                    });
                }
                db.SaveChanges();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<Workers> allWorkers;
            using (isrpoEntities db = new isrpoEntities())
            {
                allWorkers = db.Workers.ToList();
            }
            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            app.Visible = true;
            Excel.Worksheet worksheet1 = app.Worksheets.Add();
            worksheet1.Name = "Администратор";
            Excel.Worksheet worksheet2 = app.Worksheets.Add();
            worksheet2.Name = "Менеджер";
            Excel.Worksheet worksheet3 = app.Worksheets.Add();
            worksheet3.Name = "Клиент";
            worksheet1.Cells[1, 1] = "Логин";
            worksheet1.Cells[1, 2] = "Пароль";

            worksheet2.Cells[1, 1] = "Логин";
            worksheet2.Cells[1, 2] = "Пароль";

            worksheet3.Cells[1, 1] = "Логин";
            worksheet3.Cells[1, 2] = "Пароль";
            int rowindex1 = 2;
            int rowindex2 = 2;
            int rowindex3 = 2;

            foreach (var worker in allWorkers)
            {
                if (worker.RoleName == "Администратор")
                {
                    worksheet1.Cells[rowindex1, 1] = worker.LoginName;
                    worksheet1.Cells[rowindex1, 2] = GetHashString(worker.PasswordName);
                    rowindex1++;
                }
                else if (worker.RoleName == "Менеджер")
                {
                    worksheet2.Cells[rowindex2, 1] = worker.LoginName;
                    worksheet2.Cells[rowindex2, 2] = GetHashString(worker.PasswordName);
                    rowindex2++;
                }
                else if (worker.RoleName == "Клиент")
                {
                    worksheet3.Cells[rowindex3, 1] = worker.LoginName;
                    worksheet3.Cells[rowindex3, 2] = GetHashString(worker.PasswordName);
                    rowindex3++;
                }
            }
        }

        private async void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Список.json)|*.json",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            using (FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate))
            {
                List<Workers> workers = await JsonSerializer.DeserializeAsync<List<Workers>>(fs);
                using (isrpoEntities db = new isrpoEntities())
                {
                    foreach (var worker in workers)
                    {
                        db.Workers.Add(new Workers()
                        {
                            Id = worker.Id,
                            RoleName = worker.RoleName,
                            FIO = worker.FIO,
                            LoginName = worker.LoginName,
                            PasswordName = worker.PasswordName,
                        });
                    }
                    db.SaveChanges();
                }
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            List<Workers> workers = new List<Workers>();
            using (isrpoEntities db = new isrpoEntities())
            {
                workers = db.Workers.ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();

                List<Workers> category_1;
                List<Workers> category_2;
                List<Workers> category_3;

                category_1 = db.Workers.Where(x => x.RoleName == "Администратор").ToList();
                category_2 = db.Workers.Where(x => x.RoleName == "Менеджер").ToList();
                category_3 = db.Workers.Where(x => x.RoleName == "Клиент").ToList();

                var allWorkers = new List<List<Workers>>()
                {
                    category_1,
                    category_2,
                    category_3
                };
                int i = 1;
                string[] roles = { "Администраторы", "Менеджеры", "Клиенты" }; 
                foreach (var worker in allWorkers)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = roles[i - 1];
                    i++;
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table workersTable = document.Tables.Add(tableRange, worker.Count() + 1, 2);
                    workersTable.Borders.InsideLineStyle = workersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    workersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = workersTable.Cell(1, 1).Range;
                    cellRange.Text = "Логин";
                    cellRange = workersTable.Cell(1, 2).Range;
                    cellRange.Text = "Пароль";
                    workersTable.Rows[1].Range.Bold = 1;
                    workersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int j = 1;
                    foreach (var w in worker)
                    {
                        cellRange = workersTable.Cell(j + 1, 1).Range;
                        cellRange.Text = w.LoginName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = workersTable.Cell(j + 1, 2).Range;
                        cellRange.Text = GetHashString(w.PasswordName);
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        j++;
                    }
                }
                app.Visible = true;
            }
        }
    }
}
