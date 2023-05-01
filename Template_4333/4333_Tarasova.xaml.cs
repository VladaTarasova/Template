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
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Security.Cryptography;
using System.Data.Entity;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

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
        private void ImportJsonButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string jsonFilePath = openFileDialog.FileName;
                string jsonData = File.ReadAllText(jsonFilePath);
                JArray userJArray = JArray.Parse(jsonData);

                SaveUsersToDatabase(userJArray);
            }
        }
        private void SaveUsersToDatabase(JArray users)
        {
            using (workersdbEntities1 workersEntities = new workersdbEntities1())
            {
                foreach (JObject user in users)
                {
                    int userId = (int)user["Id"];
                    // Проверьте, есть ли пользователь с данным Id в БД
                    var existingUser = workersEntities.Workers.SingleOrDefault(u => u.Id == userId);

                    if (existingUser == null)
                    {
                        // Добавьте нового пользователя
                        Workers newUser = new Workers
                        {
                            Id = userId,
                            Role = (string)user["Role"],
                            FIO = (string)user["FIO"],
                            Login = (string)user["Login"],
                            Password = (string)user["Password"]
                        };

                        workersEntities.Workers.Add(newUser);
                    }
                    else
                    {
                        // Обновите данные существующего пользователя
                        existingUser.Role = (string)user["Role"];
                        existingUser.FIO = (string)user["FIO"];
                        existingUser.Login = (string)user["Login"];
                        existingUser.Password = (string)user["Password"];
                    }
                }

                workersEntities.SaveChanges();
            }
        }
        private async void ExportToWordButton_Click(object sender, RoutedEventArgs e)
        {
            var usersByPosition = await GetUsersGroupedByPosition();
            string fileName = "exported_data.docx";
            string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, fileName);
            ExportUsersToWord(usersByPosition, filePath);
            MessageBox.Show($"Данные экспортированы в файл {filePath}");
        }
        private async Task<Dictionary<string, List<Workers>>> GetUsersGroupedByPosition()
        {
            using (workersdbEntities1 workersEntities = new workersdbEntities1())
            {
                return await workersEntities.Workers
                    .GroupBy(user => user.Role)
                    .ToDictionaryAsync(group => group.Key, group => group.ToList());
            }
        }
        private void ExportUsersToWord(Dictionary<string, List<Workers>> usersByPosition, string filePath)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                // Добавить основные структуры документа
                document.AddMainDocumentPart();
                document.MainDocumentPart.Document = new Document();
                Body body = document.MainDocumentPart.Document.AppendChild(new Body());

                foreach (var positionGroup in usersByPosition)
                {
                    string position = positionGroup.Key;
                    List<Workers> workers = positionGroup.Value;

                    // Создать новый абзац с текстом заголовка
                    var headerParagraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                    var headerRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
                    var headerText = new DocumentFormat.OpenXml.Wordprocessing.Text(position);
                    headerRun.Append(headerText);
                    headerParagraph.Append(headerRun);
                    body.Append(headerParagraph);

                    // Создать таблицу
                    var table = new Table();

                    // Создать свойства таблицы
                    TableProperties tableProperties = new TableProperties(
                        new TableWidth { Type = TableWidthUnitValues.Auto, Width = "0" });

                    // Добавить свойства таблицы к таблице
                    table.Append(tableProperties);

                    // Создать и добавить строку заголовков
                    var headerRow = new TableRow();
                    var loginHeaderCell = new TableCell(new Paragraph(new Run(new Text("Логин"))));
                    var passwordHeaderCell = new TableCell(new Paragraph(new Run(new Text("Пароль"))));
                    headerRow.Append(loginHeaderCell, passwordHeaderCell);
                    table.Append(headerRow);

                    // Добавить строки данных в таблицу
                    foreach (var worker in workers)
                    {
                        var loginCell = new TableCell(new Paragraph(new Run(new Text(worker.Login))));

                        // Хэширование пароля
                        var passwordHash = ComputeSha256Hash(worker.Password);
                        var passwordCell = new TableCell(new Paragraph(new Run(new Text(passwordHash))));

                        // Создать новую строку и добавить ячейки
                        var dataRow = new TableRow();
                        dataRow.Append(loginCell, passwordCell);

                        // Добавить строку в таблицу
                        table.Append(dataRow);
                    }

                    // Добавить таблицу в тело документа
                    body.Append(table);
                    // Добавить разрыв страницы перед следующей группой данных
                    body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Break { Type = DocumentFormat.OpenXml.Wordprocessing.BreakValues.Page })));

                }

                // Сохранить изменения в документе
                document.MainDocumentPart.Document.Save();
            }
        }
        private static string ComputeSha256Hash(string rawData)
        {
            using (SHA256 sha256Hash = SHA256.Create())
            {
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(rawData));

                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    builder.Append(bytes[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }

    }
}