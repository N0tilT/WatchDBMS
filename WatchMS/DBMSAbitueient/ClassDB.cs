using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace DBMSAbitueient
{
    class ClassDB
    {
        #region Objects
        public static DataTable DataTable;
        public static OleDbConnection Connection;
        public static OleDbDataAdapter dataAdapter;
        public static DataSet DataSet;
        public static OleDbCommandBuilder ODCBuilder;
        public static string filePath = "\"Абитуриент\" Database2.mdb";
        public static string TableName = "Абитуриенты";
        public static string connectionstring = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        #endregion

        /// <summary>
        /// Открытие таблицы
        /// </summary>
        /// <param name="TableName">Имя таблицы</param>
        public static void OpenTable(string TableName)
        {
            //Устанавливаем соединение
            ClassDB.OpenConnect();

            dataAdapter = new OleDbDataAdapter("SELECT * FROM " + $"[{TableName}]", Connection);

            ClassDB.TableName = TableName;

            //Заполняем DataSet с помощью dataAdapter
            ODCBuilder = new OleDbCommandBuilder(dataAdapter);
            DataSet = new DataSet();
            try
            {
                dataAdapter.Fill(DataSet);
                DataTable = DataSet.Tables[0];
            }
            catch (Exception)
            {
                return;
            }
        }

        /// <summary>
        /// Получение списка имён таблиц базы данных из метаданных
        /// </summary>
        /// <returns></returns>
        public static List<string> GetTablesNames()
        {
            if (Connection.State == 0)
                OpenConnect();

            DataTable dbTables = Connection.GetSchema("Tables", new[] { null, null, null, "TABLE" });

            List<String> TableNameList = new List<string>();
            TableNameList.AddRange(
                from DataRow item in dbTables.Rows 
                select item[2].ToString());

            return TableNameList;
        }

        /// <summary>
        /// Получение списка имён отображений базы данных из метаданных
        /// </summary>
        /// <returns></returns>
        public static List<string> GetViewsNames()
        {
            OpenConnect();
            if (Connection.State == 0)
                OpenConnect(); 

            DataTable dbViews = ClassDB.Connection.GetSchema("Views");
            List<String> ViewsNameList = new List<string>();

            foreach (DataRow dataRow in dbViews.Rows)
            {
                ViewsNameList.Add(dataRow[2].ToString());
                ViewsNameList.Add(dataRow[3].ToString());
            }
            return ViewsNameList;
        }
        public static void OpenConnect()
        {
            Connection = new OleDbConnection(connectionstring);
            Connection.Open();
        }
        public static void CloseConnect()
        {
            Connection.Dispose();
        }

        /// <summary>
        /// Запрос файла у пользователя
        /// </summary>
        public static void GetFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())    //Открытие файлового диалога
            {
                openFileDialog.InitialDirectory = "с:\\";
                openFileDialog.Filter = "accdb files (*.accdb)|*.accdb|db files (*.mdb)|*.mdb|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Получаем путь и меняем его в строке соединения
                    filePath = openFileDialog.FileName;
                    connectionstring = connectionstring.Replace(
                        connectionstring.Substring(
                            connectionstring.IndexOf("Data Source=") + "Data Source=".Length,
                            connectionstring.Length - connectionstring.IndexOf("Data Source=") - "Data Source=".Length), 
                        filePath);
                }
            }

        }


        /// <summary>
        /// Сохранение данных
        /// </summary>
        public static void SaveData(string TableName)
        {
            OleDbCommandBuilder cmdb = new OleDbCommandBuilder(dataAdapter);

            //Заключение параметров в [] для обеспечения читаемости
            cmdb.QuotePrefix = "[";
            cmdb.QuoteSuffix = "]";

            dataAdapter.UpdateCommand = cmdb.GetUpdateCommand();
            dataAdapter.Update(DataSet);
        }

        /// <summary>
        /// Создание запроса
        /// </summary>
        /// <param name="Command">Строка команды</param>
        public static void CreateQuery(string Command)
        {
            dataAdapter = new OleDbDataAdapter(Command, Connection);
            DataSet = new DataSet();
            dataAdapter.Fill(DataSet);
        }

        /// <summary>
        /// Формирование отчета в Word
        /// </summary>
        /// <param name="DataTable">Данные для отчета</param>
        public static void WriteWordReport(DataTable DataTable)
        {
            // Инициализация нового документа
            Word.Document doc;
            Word.Application app = new Word.Application
            {
                Caption = "Отчёт",
                Visible = true
            };
            doc = app.Documents.Add();

            // Указание даты создания отчёта
            Word.Range nRange = doc.Range(0, 0);
            nRange.Text = DateTime.Now.ToString();

            // Создание таблицы на документе Word
            Word.Range WRange = doc.Range(0, 0);
            Word.Table WTable = doc.Tables.Add(WRange, DataTable.Rows.Count + 1, DataTable.Columns.Count, 
                Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitWindow);

            // Заполнение таблицы данными из отчёта
            WTable = doc.Tables[1];
            Word.Range WCelRange = doc.Tables[1].Range;
            for (int i = 0; i < WTable.Columns.Count; i++)
            {
                WCelRange = WTable.Cell(1, i + 1).Range;
                WCelRange.Text = DataTable.Columns[i].Caption;
            }

            for (int i = 1; i <= DataTable.Rows.Count; i++)
            {
                for (int j = 0; j < WTable.Columns.Count; j++)
                {
                    WCelRange = WTable.Cell(i + 1, j + 1).Range;
                    WCelRange.Text = DataTable.Rows[i - 1][j].ToString();
                }
            }
        }
        
        public static void WriteExcelReport(DataTable DT)
        {
            // Создание нового листа Excel
            Excel.Application app = new Excel.Application()
            {
                Caption = "Отчёт"
            };
            Excel.Workbook book = app.Workbooks.Add();
            Excel.Worksheet sheet = book.Worksheets.Item[1];

            app.Visible = true;
            for (int i = 0; i < DT.Columns.Count; i++)
            {
                sheet.Cells[1, i+1] = DT.Columns[i].Caption;
            }
            
            for (int i = 0; i < DT.Rows.Count ; i++)
            {
                for (int j = 0; j < DT.Columns.Count; j++)
                {
                    sheet.Cells[j + 1][i+2].Value = DT.Rows[i][j];
                }
            }
            sheet.Columns.AutoFit();
        }
    }
}
