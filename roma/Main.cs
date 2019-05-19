using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Drawing.Printing;

namespace roma
{
    public partial class Main : Form
    {
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=бд.mdb;";
        public bool Dgb;
        private OleDbConnection myConnection;
        reg r1 = new reg();
        reg regf = new reg();

        public Main()
        {

            
            InitializeComponent();
            // создаем экземпляр класса OleDbConnection
            myConnection = new OleDbConnection(connectString);

            // открываем соединение с БД
            myConnection.Open();
            
            r1.Hide();
            regf.ShowDialog();
            tabControl1.BackColor = Color.Red;
            
                


        }
        
        private void dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                e.Value = e.RowIndex + 1;
            }
        }
        public void Main_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableSF". При необходимости она может быть перемещена или удалена.
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableAF". При необходимости она может быть перемещена или удалена.
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableNF". При необходимости она может быть перемещена или удалена.
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableAF". При необходимости она может быть перемещена или удалена.
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableNF". При необходимости она может быть перемещена или удалена.
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);





        }

        private void button1_Click(object sender, EventArgs e)
        {
                PrintDialog printDialog = new PrintDialog();
                PrintDocument printText = new PrintDocument();
                printText.DocumentName = "ss";
                printDialog.Document = printText;
                printDialog.AllowSelection = true;
                printDialog.AllowSomePages = true;

                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    printText.PrintPage += new PrintPageEventHandler(prin);
                    printText.Print();
                }     
        }
        void prin(object sender, PrintPageEventArgs e)
        {
            Graphics grap = e.Graphics;
            grap.DrawString(textBox1.Text, Font, new SolidBrush(Color.Black), 0, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            int a = dataGridView1.CurrentRow.Index;
            string b1 = dataGridView1.Rows[a].Cells[1].Value.ToString();
            string b2 = dataGridView1.Rows[a].Cells[2].Value.ToString();
            string b3 = dataGridView1.Rows[a].Cells[3].Value.ToString();
            //string b = "green";
            //
            //dataGridView1.Rows[a].Cells["bdcolorDataGridViewTextBoxColumn"].Value = b;
            //dataGridView1.Rows[a].Cells[4].Value = textBox1.Text;
            //this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            // текст запроса
            string query = "INSERT INTO TableAF(Заявитель,Проблема,Дата)"
                                         + "VALUES ('" + b1 + "','" + b2 + "','" + b3 + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            dataGridView1.Rows.Remove(dataGridView1.Rows[a]);            
            MessageBox.Show("Статус заявки переведен на 'выполнена'. Заявка отправлена в архив. Дождитесь открытия отчета Word");
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string path = Environment.CurrentDirectory + "/1.dot";
            //Загружаем документ
            Microsoft.Office.Interop.Word.Document doc = null;
            object fileName = path;
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Теперь у нас есть документ который мы будем менять.

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //Задаём параметры замены и выполняем замену.
            string b4 = dateTimePicker1.Text;
            object findText = "<date>";
            object replaceWith = b4;
            object replace = 2;
            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);

            ///////////////////
            ///
            object findText1 = "<z>";
            object replaceWith1 = b1;
            object replace1 = 2;
            app.Selection.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1,
            ref replace1, ref missing, ref missing, ref missing, ref missing);
            //////////////////
            ///
            object findText2 = "<t>";
            object replaceWith2 = b2;
            object replace2 = 2;
            app.Selection.Find.Execute(ref findText2, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith2,
            ref replace2, ref missing, ref missing, ref missing, ref missing);  
            
            ///////////////////////////
            
            object findText4 = "<zt>";
            object replaceWith4 = b3;
            object replace4 = 2;
            app.Selection.Find.Execute(ref findText4, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith4,
            ref replace4, ref missing, ref missing, ref missing, ref missing);            

            // Открываем документ для просмотра.
             app.Visible = true;
            
        }

        private void tableNFBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void tableNFBindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void Main_FormClosed(object sender, FormClosedEventArgs e)
        {
            myConnection.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int a = dataGridView2.CurrentRow.Index;
            string b1 = dataGridView2.Rows[a].Cells[1].Value.ToString();
            string b2 = dataGridView2.Rows[a].Cells[2].Value.ToString();
            string b3 = dataGridView2.Rows[a].Cells[3].Value.ToString();            
            string query = "INSERT INTO TableNF(Заявитель,Проблема,Дата)"
                                         + "VALUES ('" + b1 + "','" + b2 + "','" + b3 + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            dataGridView2.Rows.Remove(dataGridView2.Rows[a]);
            MessageBox.Show("Заявка добавлена в список");
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (dialog_delete fMyForm = new dialog_delete())
            {
                if (fMyForm.ShowDialog() == DialogResult.OK)
                {
                    int a = dataGridView1.CurrentRow.Index;
                    dataGridView1.Rows.Remove(dataGridView1.Rows[a]);
                    dataGridView1.Refresh();
                    this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
                    this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
                    this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
                    this.бдDataSet.Clear();
                    this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
                    this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
                    this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
                }

            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void опрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("СПО МБОУ СОШ№89 Программа учета заявок обслуживания компьютерной техники и устройств. 2019. Дипломный проект. ©Копытов Роман Владимирович - Красноярский колледж радиоэлектроники и информационных технологий - ПИ-3.16К");
        }

        private void сменитьПользователяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            reg regf = new reg();
            regf.ShowDialog();
        }

        private void внестиЗаявкуToolStripMenuItem_Click(object sender, EventArgs e)
        {           
            reg regf = new reg();
            regf.ShowDialog();
        }

        private void изменитьСтатусНаВОжиданииToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            int a = dataGridView1.CurrentRow.Index;
            string b1 = dataGridView1.Rows[a].Cells[1].Value.ToString();
            string b2 = dataGridView1.Rows[a].Cells[2].Value.ToString();
            string b3 = dataGridView1.Rows[a].Cells[3].Value.ToString();
            
            string query = "INSERT INTO TableSF(Заявитель,Проблема,Дата)"
                                         + "VALUES ('" + b1 + "','" + b2 + "','" + b3 + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            dataGridView1.Rows.Remove(dataGridView1.Rows[a]);
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            MessageBox.Show("Статус заявки переведен на 'Выполняется'");

            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int a = dataGridView3.CurrentRow.Index;
            string b1 = dataGridView3.Rows[a].Cells[1].Value.ToString();
            string b2 = dataGridView3.Rows[a].Cells[2].Value.ToString();
            string b3 = dataGridView3.Rows[a].Cells[3].Value.ToString();
            string query = "INSERT INTO TableNF(Заявитель,Проблема,Дата)"
                                         + "VALUES ('" + b1 + "','" + b2 + "','" + b3 + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            dataGridView3.Rows.Remove(dataGridView3.Rows[a]);
            MessageBox.Show("Заявка добавлена в список");
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {
            
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {
            
        }

        private void tabPage1_Enter(object sender, EventArgs e)
        {
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
        }

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
            this.бдDataSet.Clear();
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
        }

        private void tabPage2_Enter(object sender, EventArgs e)
        {
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.бдDataSet.Clear();
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            using (dialog_delete fMyForm = new dialog_delete())
            {
                if (fMyForm.ShowDialog() == DialogResult.OK)
                {
                    int a = dataGridView3.CurrentRow.Index;
                    dataGridView3.Rows.Remove(dataGridView3.Rows[a]);
                    dataGridView3.Refresh();
                    this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
                    this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
                    this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
                    this.бдDataSet.Clear();
                    this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
                    this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
                    this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
                }

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

            int a = dataGridView3.CurrentRow.Index;
            string b1 = dataGridView3.Rows[a].Cells[1].Value.ToString();
            string b2 = dataGridView3.Rows[a].Cells[2].Value.ToString();
            string b3 = dataGridView3.Rows[a].Cells[3].Value.ToString();
            //string b = "green";
            //
            //dataGridView3.Rows[a].Cells["bdcolorDataGridViewTextBoxColumn"].Value = b;
            //dataGridView3.Rows[a].Cells[4].Value = textBox1.Text;
            //this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            // текст запроса
            string query = "INSERT INTO TableAF(Заявитель,Проблема,Дата)"
                                         + "VALUES ('" + b1 + "','" + b2 + "','" + b3 + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            dataGridView3.Rows.Remove(dataGridView3.Rows[a]);
            MessageBox.Show("Статус заявки переведен на 'выполнена'. Заявка отправлена в архив. Дождитесь открытия отчета Word");
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string path = Environment.CurrentDirectory + "/1.dot";
            //Загружаем документ
            Microsoft.Office.Interop.Word.Document doc = null;
            object fileName = path;
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Теперь у нас есть документ который мы будем менять.

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //Задаём параметры замены и выполняем замену.
            string b4 = dateTimePicker1.Text;
            object findText = "<date>";
            object replaceWith = b4;
            object replace = 2;
            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);

            ///////////////////
            ///
            object findText1 = "<z>";
            object replaceWith1 = b1;
            object replace1 = 2;
            app.Selection.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1,
            ref replace1, ref missing, ref missing, ref missing, ref missing);
            //////////////////
            ///
            object findText2 = "<t>";
            object replaceWith2 = b2;
            object replace2 = 2;
            app.Selection.Find.Execute(ref findText2, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith2,
            ref replace2, ref missing, ref missing, ref missing, ref missing);

            ///////////////////////////

            object findText4 = "<zt>";
            object replaceWith4 = b3;
            object replace4 = 2;
            app.Selection.Find.Execute(ref findText4, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith4,
            ref replace4, ref missing, ref missing, ref missing, ref missing);

            // Открываем документ для просмотра.
            app.Visible = true;
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
        }

        private void button9_Click_2(object sender, EventArgs e)
        {
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Update(this.бдDataSet.TableSF);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            this.бдDataSet.Clear();
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableSFTableAdapter.Fill(this.бдDataSet.TableSF);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MessageBox.Show("СПО МБОУ СОШ№89 Программа учета заявок обслуживания компьютерной техники и устройств. 2019. Дипломный проект. ©Копытов Роман Владимирович - Красноярский колледж радиоэлектроники и информационных технологий - ПИ-3.16К");

        }

        
    }
}