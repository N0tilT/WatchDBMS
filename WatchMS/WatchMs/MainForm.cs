using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace DBMSAbitueient
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //Создание "кнопок" на форме
            //Обьединение событий нажатия на иконку кнопки, её название и область
            groupBoxTables.Click += new EventHandler(labelTables_Click);
            groupBoxQuery.Click += new EventHandler(labelQuery_Click);
            groupBoxInfo.Click += new EventHandler(labelInfo_Click);
            groupBoxSettings.Click += new EventHandler(labelSettings_Click);
            groupBoxExit.Click += new EventHandler(labelExit_Click);

            pictureBoxTables.Click += new EventHandler(labelTables_Click);
            pictureBoxQuery.Click += new EventHandler(labelQuery_Click);
            pictureBoxInfo.Click += new EventHandler(labelInfo_Click);
            pictureBoxSettings.Click += new EventHandler(labelSettings_Click);
            pictureBoxExit.Click += new EventHandler(labelExit_Click);
        }

        #region Objects
        List<string> ViewsNames = new List<string>();      //[View_Name, SQL_Query]
        List<string> Queries = new List<string>();         //SQL - запросы
        List<string> QueriesNames = new List<string>();    //Название SQL - запросов
        BindingSource BDS;
        #endregion

        #region ListFunctions
        /// <summary>
        /// Разделение массива ViewsNames[Name,Query,Name,Query] на Queries[Query,Query] и QueriesNames[Name,Name]
        /// </summary>
        private void SplitViewsNames()
        {
            for (int i = 0; i < ViewsNames.Count; i++)
            {
                //Чётные элементы - названия запросов, нечётные - их текст
                if (i % 2 == 1) Queries.Add(ViewsNames[i]);
                else QueriesNames.Add(ViewsNames[i]);
            }
        }

        /// <summary>
        /// Преобразование SQL-запроса ACCESS в запрос, читаемый С#
        /// Замены в конструкции WHERE:
        /// " -> '
        /// * -> '
        /// Like -> LIKE
        /// </summary>
        private void MakeQueriesReadable()
        {
            for (int i = 0; i < Queries.Count; i++)
            {
                //Просматриваем запросы. Если есть WHERE, заменяем
                string query = Queries[i];
                if (query.IndexOf("WHERE") != -1)
                {
                    //то, что нужно заменить
                    string NeedToReplace = query.Substring(query.IndexOf("WHERE"), query.Length - query.IndexOf("WHERE"));  //WHERE...

                    //то, чем заменяем (редактируем NeedToReplace)
                    string WhatToReplace = NeedToReplace.Replace('\"', '\'');
                    WhatToReplace = WhatToReplace.Replace('*', '%');
                    WhatToReplace = WhatToReplace.Replace("Like", "LIKE");

                    //Заменяем в запросе
                    query = query.Replace(NeedToReplace, WhatToReplace);
                }

                //Меняем элемент в контейнере
                Queries[i] = query;
            }
        }
        #endregion

        #region Menu
        /// <summary>
        /// Сворачивание меню
        /// </summary>
        private void pictureBoxMenuSlide_Click(object sender, EventArgs e)
        {
            CollapseMenu();
        }

        /// <summary>
        /// Сворачивание меню
        /// </summary>
        private void CollapseMenu()
        {
            if (this.panelMenu.Width > 55) //Уменьшение меню
            {
                panelMenu.Width = 55;
                menuStripSettings.Location = new Point { X = 55, Y = 374 };
                labelHeader.Visible = false;
                pictureBoxMenuSlide.Dock = DockStyle.Top;
            }
            else  //Увеличение меню
            {
                panelMenu.Width = 250;
                menuStripSettings.Location = new Point { X = 250, Y = 374 };
                labelHeader.Visible = true;
                pictureBoxMenuSlide.Dock = DockStyle.None;
            }
        }


        /// <summary>
        /// О программе
        /// </summary>
        private void labelInfo_Click(object sender, EventArgs e)
        {
            UpdateMenu();
            groupBoxInfo.BackColor = Color.FromArgb(92, 114, 158);
            MessageBox.Show("   Программа представляет собой систему управления базой данных. \n" +
                "Она позволяет просматривать и редактировать существующие таблицы, отчёты и запросы к базе данных. \n" +
                "База данных по умолчанию - БД \"Абитуриент\", но пользователь может загрузить и собственную базу данных с помощью диалогового окна.\n", "О программе");

        }

        /// <summary>
        /// Настройки
        /// </summary>
        private void labelSettings_Click(object sender, EventArgs e)
        {
            UpdateMenu();

            //Окрашивание в цвет нажатия
            groupBoxSettings.BackColor = Color.FromArgb(92, 114, 158);

            //Отображение меню настроек
            if (menuStripSettings.Visible) menuStripSettings.Visible = false;
            else menuStripSettings.Visible = true;
        }

        /// <summary>
        /// Выход
        /// </summary>
        private void labelExit_Click(object sender, EventArgs e)
        {
            UpdateMenu();

            //Окрашивание в цвет нажатия
            groupBoxExit.BackColor = Color.FromArgb(92, 114, 158);

            Application.Exit();
        }

        /// <summary>
        /// Окрашивание всех кнопок, окрашеных в цвет нажатия, в обычный цвет
        /// </summary>
        private void UpdateMenu()
        {
            foreach (Control gbMenu in groupBoxMenu.Controls)
            {
                if (gbMenu.BackColor == Color.FromArgb(92, 114, 158))
                    gbMenu.BackColor = Color.FromArgb(140, 163, 211);
            }
        }

        #endregion

        #region Connection
        /// <summary>
        /// Подключение БД, пользователь указывает файл. По умолчанию - "Абитуриент"
        /// После подключения БД включается остальные пункты меню
        /// </summary>
        private void подключениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Получение файла через файловый диалог
            try
            {
                ClassDB.GetFile();
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message);
            }

            //Прячем меню
            menuStripSettings.Visible = false;

            labelState.BackColor = Color.Orange;
            panelState.BackColor = Color.Orange;
            labelState.Text = "Состояние: Подключение не выполнено";

            try
            {
                //Попытка установить соединение
                ClassDB.OpenConnect();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение\n" + $"Connection string - {ClassDB.connectionstring}", "Ошибка");

                return;
            }

            MessageBox.Show($"Подключение с базой данных {ClassDB.filePath} успешно установлено", "Успех!");

            //Включаем остальные пункты меню
            EnableMenu();
            //Обновляем интерфейс
            UpdateUI();

            //Обновляем объекты с запросами
            ViewsNames = new List<string>();
            Queries = new List<string>();
            QueriesNames = new List<string>();

            //Получаем список запросов
            ViewsNames = ClassDB.GetViewsNames();
            SplitViewsNames();
            MakeQueriesReadable();
        }

        private void pictureBoxRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                //Попытка установить соединение
                ClassDB.OpenConnect();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение\n" + $"Connection string - {ClassDB.connectionstring}", "Ошибка");

                return;
            }
        }

        /// <summary>
        /// Обновление интерфейса
        /// </summary>
        private void UpdateUI()
        {
            //Обнуление всех текстовых полей, источников данных для элементов управления
            comboBoxMain.Text = "";
            dataGridViewMain.DataSource = null;
            textBoxTableInfo.Text = "";
            labelInfoTables.Text = "";

            //Скрытие всех элементов управления
            panel3.Visible = false;
            tableLayoutPanel1.Visible = false;

            //Выключение всех кнопок на рабочей области
            buttonSave.Enabled = false;
            buttonDelete.Enabled = false;
            buttonAdd.Enabled = false;
            buttonRunQuery.Enabled = false;
            buttonEdit.Enabled = false;
            buttonPrintWord.Enabled = false;
        }

        /// <summary>
        /// Включение пунктов меню - Таблицы и Запросы
        /// </summary>
        private void EnableMenu()
        {
            groupBoxQuery.Enabled = true;
            groupBoxTables.Enabled = true;
            labelState.BackColor = Color.Green;
            panelState.BackColor = Color.Green;
            labelState.Text = "Состояние: Успешное подключение - " + ClassDB.filePath;
        }

        #endregion

        #region Tables
        /// <summary>
        /// Открытие таблиц
        /// </summary>
        private void labelTables_Click(object sender, EventArgs e)
        {
            UpdateMenu();

            //Окрашивание кнопки в цвет нажатия
            groupBoxTables.BackColor = Color.FromArgb(92, 114, 158);

            //Включаем элементы управления таблицами
            EnableUITables();

            //Обновляем список таблиц
            UpdateCBTables();
        }

        /// <summary>
        /// Включение элементов управления таблицами
        /// </summary>
        private void EnableUITables()
        {
            //Открываем рабочую область
            panel3.Visible = true;
            tableLayoutPanel1.Height = 125;
            tableLayoutPanel1.Visible = true;
            textBoxTableInfo.Height = 79;

            //Заполняем текстовые поля
            comboBoxMain.Text = "Выберите таблицу";
            labelMainHeader.Text = "Таблицы";
            labelInfoTables.Text = "Информация о таблице";

            //Включаем элементы управления таблицами
            comboBoxMain.Enabled = true;
            buttonSave.Visible = true;
            buttonDelete.Visible = true;
            buttonAdd.Visible = true;
            buttonSave.Enabled = true;
            buttonDelete.Enabled = true;
            buttonAdd.Enabled = true;

            //Скрываем элементы управления запросами
            buttonPrintWord.Visible = false;
            buttonPrintExcel.Visible = false;
            buttonRunQuery.Visible = false;
            buttonEdit.Visible = false;
        }

        /// <summary>
        /// Обновление ComboBox со списком таблиц
        /// </summary>
        private void UpdateCBTables()
        {
            //Получаем имена всех таблиц из базы данных
            List<string> TablesNames = ClassDB.GetTablesNames();    //[Table_Name,Table_Name2...]

            //Заносим таблицы в ComboBox
            comboBoxMain.DataSource = TablesNames;

            //Подготавливаем рабочую область к работе с таблицами
            comboBoxMain.Text = "Выберите таблицу";
            dataGridViewMain.DataSource = null;
            textBoxTableInfo.Text = "";
        }

        /// <summary>
        /// Сохранение изменений в таблице
        /// </summary>
        private void buttonSave_Click(object sender, EventArgs e)
        {
            //В случае, когда сохранять нечего
            if (ClassDB.DataSet == null)
                return;
            else
            {
                DialogResult result = MessageBox.Show("Вы хотите сохранить изменения в базе данных?",
                "Сохранение", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                    try
                    {
                        //Попытка сохранения данных
                        BDS.EndEdit();
                        ClassDB.SaveData(comboBoxMain.Text);
                    }
                    catch (Exception q)
                    {
                        MessageBox.Show(q.Message);
                        return;
                    }

            }

            dataGridViewMain.Columns[0].ReadOnly = true;
        }

        /// <summary>
        /// Добавление новой строки
        /// </summary>
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            DataRow row = ClassDB.DataSet.Tables[0].NewRow();
            ClassDB.DataSet.Tables[0].Rows.Add(row);
            dataGridViewMain.Columns[0].ReadOnly = false;
        }

        /// <summary>
        /// Удаление выбранных строк
        /// </summary>
        private void buttonDelete_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewMain.SelectedRows)
            {
                dataGridViewMain.Rows.Remove(row);
            }
        }

        #endregion

        #region Queries
        /// <summary>
        /// Открытие запросов
        /// </summary>
        private void labelQuery_Click(object sender, EventArgs e)
        {
            UpdateMenu();

            //Окрашивание кнопки в цвет нажатия
            groupBoxQuery.BackColor = Color.FromArgb(92, 114, 158);

            //Включаем элементы управления запросами
            EnableUIQuery();

            //Обновляем список запросов
            UpdateCBQuery();

        }

        /// <summary>
        /// Включение элементов управления запросами
        /// </summary>
        private void EnableUIQuery()
        {
            //Открываем рабочей области
            panel3.Visible = true;
            tableLayoutPanel1.Visible = true;
            tableLayoutPanel1.Height = 165;
            textBoxTableInfo.Height = 117;

            //Заполняем необходимые текстовые поля
            comboBoxMain.Text = "Выберите запрос";
            labelMainHeader.Text = "Запросы";
            labelInfoTables.Text = "SQL-запрос:";

            //Включаем элементы управления запросами
            comboBoxMain.Enabled = true;
            buttonPrintWord.Visible = true;
            buttonRunQuery.Visible = true;
            buttonPrintExcel.Visible = true;
            buttonEdit.Visible = true;
            buttonRunQuery.Enabled = true;
            buttonEdit.Enabled = true;
            buttonPrintWord.Enabled = true;
            buttonPrintExcel.Enabled = true;

            //Скрываем элементы управления таблицами
            buttonSave.Visible = false;
            buttonDelete.Visible = false;
            buttonAdd.Visible = false;
        }

        /// <summary>
        /// Обновление ComboBox - списка запросов
        /// </summary>
        private void UpdateCBQuery()
        {
            //Заполнение списка запросов
            comboBoxMain.DataSource = QueriesNames;

            //Подготовка рабочей области к работе с запросами
            comboBoxMain.Text = "Выберите запрос";
            dataGridViewMain.DataSource = null;
            textBoxTableInfo.Text = "";
        }

        /// <summary>
        /// Выполнение запроса из текстового поля
        /// </summary>
        private void buttonRunQuery_Click(object sender, EventArgs e)
        {
            try
            {
                //Выполнение запроса
                ClassDB.CreateQuery(textBoxTableInfo.Text);

                //Заполнение рабочей области результатом запроса
                dataGridViewMain.DataSource = ClassDB.DataSet.Tables[0].DefaultView;
                //Выравнивание ширины столбцов
                dataGridViewMain.AutoResizeColumns();

                //Привязка элементов управления формой к данным
                BDS = new BindingSource { DataSource = dataGridViewMain.DataSource };
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message, "Ошибка");
                return;
            }
        }

        /// <summary>
        /// Редактирование запроса в текстовом поле
        /// </summary>
        private void buttonEdit_Click(object sender, EventArgs e)
        {
            textBoxTableInfo.ReadOnly = false;

            comboBoxMain.Text = "Запрос пользователя";
        }

        /// <summary>
        /// Создание отчёта по запросу в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonPrintWord_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.DataSource == null)
                return;
            try
            {
                //Создание отчёта в Word
                ClassDB.WriteWordReport(ClassDB.DataSet.Tables[0]);
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message, "Ошибка");
            }
        }
        
        /// <summary>
         /// Создание отчёта по запросу в Excel
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
        private void buttonPrintExcel_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.DataSource == null)
                return;
            try
            {
                //Создание отчёта в Excel
                ClassDB.WriteExcelReport(ClassDB.DataSet.Tables[0]);
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message, "Ошибка");
            }
        }
        #endregion

        #region ComboBox
        /// <summary>
        /// Открытие выбранной пользователем таблицы/запроса
        /// </summary>
        private void comboBoxtTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Запрет на редактирование текстового поля с информацией
            textBoxTableInfo.ReadOnly = true;

            //Индекс выбранного запроса
            int index = QueriesNames.IndexOf(comboBoxMain.Text);

            if (groupBoxQuery.BackColor == Color.FromArgb(92, 114, 158))    //Выбраны запросы
            {
                //Вывод информации о запросе
                textBoxTableInfo.Text = Queries[index];
                textBoxTableInfo.ScrollBars = ScrollBars.Vertical;

                //Выполнение запроса
                ClassDB.CreateQuery(textBoxTableInfo.Text);

                //Заполнение рабочей области результатом запроса
                dataGridViewMain.DataSource = ClassDB.DataSet.Tables[0].DefaultView;
                //Выравнивание ширины столбцов
                dataGridViewMain.AutoResizeColumns();

                //Привязка элементов управления формой к данным
                BDS = new BindingSource { DataSource = dataGridViewMain.DataSource };

            }
            else  //Выбраны таблицы
            {
                //Открываем нужную таблицу
                ClassDB.OpenTable(comboBoxMain.Text);
                if (ClassDB.DataTable != null)
                {
                    //Заполняем рабочую область
                    dataGridViewMain.DataSource = ClassDB.DataTable;

                    //Выравниваем ширину столбцов
                    dataGridViewMain.AutoResizeColumns();

                    //Запрет на редактирование ключевого поля
                    dataGridViewMain.Columns[0].ReadOnly = true;

                    //Привязка элементов управления формой к данным
                    BDS = new BindingSource { DataSource = dataGridViewMain.DataSource };
                }

                //Вывод информации о таблице
                textBoxTableInfo.Text = $"Количество строк - {ClassDB.DataTable.Rows.Count}; \r\n" +
                $"Количество столбцов - {ClassDB.DataTable.Columns.Count};";
                textBoxTableInfo.ScrollBars = ScrollBars.Horizontal;
            }

        }
        #endregion
    }
}
