using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace roma
{
    public partial class reg : Form
    {
        
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=бд.mdb;";
        private OleDbConnection myConnection;
        public reg()
        {
            InitializeComponent();
            // создаем экземпляр класса OleDbConnection
            myConnection = new OleDbConnection(connectString);

            // открываем соединение с БД
            myConnection.Open();


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            myConnection.Close();
            if (textBox1.Text == "admin")
            {
                if (textBox2.Text == "12345")
                {
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не верный пароль");
                }
            }
            if (textBox1.Text == "qq")
            {
                if (textBox2.Text == "qq")
                {
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не верный пароль");
                }

            }
            
            if (textBox1.Text == null | textBox2.Text == null)
            {
                MessageBox.Show("Не верный логин либо пароль");
            }

            

        }
        private void reg_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (textBox1.Text != "admin" | textBox2.Text != "12345")
                if (textBox1.Text != "qq" | textBox2.Text != "qq")
                {
                    Application.Exit();

                }

            myConnection.Close();
        }
         

        private void reg_Load(object sender, EventArgs e)
        {
            panel1.Show();
            panel2.Hide();
            panel3.Hide();
        }

        public void button2_Click(object sender, EventArgs e)
        {
            panel3.Show();
            panel2.Hide();
            panel1.Hide();

        }

        public void button3_Click(object sender, EventArgs e)
        {
            panel1.Hide();
            panel2.Show();
            panel3.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel1.Show();
            panel2.Hide();
            panel3.Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel1.Show();
            panel2.Hide();
            panel3.Hide();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public void button5_Click(object sender, EventArgs e)
        {
            string name = richTextBox1.Text;
            string trouble = richTextBox2.Text;
            string date = dateTimePicker1.Text;
            // cmd = new SqlCommand("Insert Into into Values ('" + richTextBox1.Text + "', '" + richTextBox2.Text + "', '" + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "')", conn); // создаем SQL запрос
            // текст запроса
            string query = "INSERT INTO TableNF(Заявитель,Проблема,Дата)"
                                         + "VALUES ('" + name + "','" + trouble + "','" + date + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            MessageBox.Show("Данные переданны");
            richTextBox1.Clear();
            richTextBox2.Clear();
            dateTimePicker1.ResetText();
            panel1.Show();
            panel2.Hide();
            panel3.Hide();
        }
    }
    

}
