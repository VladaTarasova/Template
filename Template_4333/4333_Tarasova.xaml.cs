using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.Windows.Shapes;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Tarasova.xaml
    /// </summary>
    public partial class _4333_Tarasova : Window
    {
        public _4333_Tarasova()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new
            Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
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

            using (workersdbEntities1 workersEntities = new workersdbEntities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (string.IsNullOrEmpty(list[i, 0]))
                    {
                        continue;
                    }
                    workersEntities.Workers.Add(new Workers()
                    {
                        Role = list[i, 0],
                        FIO = list[i, 1],
                        Login = list[i, 2],
                        Password = list[i, 3]
                    });
                }
                workersEntities.SaveChanges();
            }
        }
        private void BnExport_Click(object sender, RoutedEventArgs e)
        {

            Dictionary<string, List<Workers>> ByData = new Dictionary<string, List<Workers>>();
            using (workersdbEntities1 workersEntities = new workersdbEntities1())
            {
                var allWorkers = workersEntities.Workers.ToList().GroupBy(w => w.Role);

                foreach (var group in allWorkers)
                {
                    ByData[group.Key] = group.ToList();
                }
            }
            var app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            app.Visible = true;

            foreach (var worker in ByData)
            {
                string role = worker.Key;
                List<Workers> workers = worker.Value;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = app.Worksheets.Add();

                if (role != "")
                {
                    worksheet.Name = role;

                    worksheet.Cells[1, 1] = "Логин";
                    worksheet.Cells[1, 2] = "Пароль";
                }
                int rowIndex = 2;

                foreach (Workers work in workers)
                {
                    worksheet.Cells[rowIndex, 1] = work.Login;
                    worksheet.Cells[rowIndex, 2] = work.Password;
                    rowIndex++;
                }
            }
        }
    }
}
