using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private string _connectionString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|Worker.mdb;";
        private OleDbConnection _connection;


        public Form1()
        {
            InitializeComponent();
            DataGridViewDbLoad("tableName");
        }


        /// <summary>
        /// Пример вывода таблицы в DataGridView
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        private void DataGridViewDbLoad(string tableName)
        {
            _connection = new OleDbConnection(_connectionString);

            _connection.Open();

            // виртуальная таблица
            DataTable table = new DataTable();

            // SQL-запрос
            string command = $"SELECT * FROM [{tableName}]";

            // команда для базы данных
            OleDbDataAdapter adapter = new OleDbDataAdapter(command, _connection);
           
            // заполнение виртуальной таблицы
            adapter.Fill(table);

            // передача виртуальной таблицы в качестве источника данных
            dataGridView1.DataSource = table;

            _connection.Close();
        }

        private void button1_Click()
        {
            
        }

        /// <summary>
        /// Метод добавляет новую запись в таблицу Категории
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInsert_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(textBox1.Text);
            string type = textBox2.Text;
            int discount = Convert.ToInt32(textBox3.Text);

            // запрос
            string query = $"INSERT INTO [tableName] " +
                $"([ID], [Категория товара], [Скидка %]) " +
                $"VALUES ({id}, '{type}', {discount})";

            _connection = new OleDbConnection(_connectionString);

            _connection.Open();

            OleDbCommand command = new OleDbCommand(query, _connection);

            // запрос, который не выводит ничего, а только выполняет команду
            command.ExecuteNonQuery();

            _connection.Close();
            DataGridViewDbLoad("tableName");
        }

        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(textBox1.Text);
            string type = textBox2.Text;
            int discount = Convert.ToInt32(textBox3.Text);

            // запрос
            string query = $"UPDATE [tableName] SET [Категория товара] = '{type}', " +
                $"[Скидка %] = {discount} " +
                $"WHERE [ID] = {id}";

            _connection = new OleDbConnection(_connectionString);

            _connection.Open();

            OleDbCommand command = new OleDbCommand(query, _connection);

            // запрос, который не выводит ничего, а только выполняет команду
            command.ExecuteNonQuery();

            _connection.Close();
            DataGridViewDbLoad("tableName");
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(textBox1.Text);            

            // запрос
            string query = $"DELETE FROM [tableName] " +
                $"WHERE [ID] = {id}";

            _connection = new OleDbConnection(_connectionString);

            _connection.Open();

            OleDbCommand command = new OleDbCommand(query, _connection);

            // запрос, который не выводит ничего, а только выполняет команду
            command.ExecuteNonQuery();

            _connection.Close();
            DataGridViewDbLoad("tableName");
        }
    }
}
