using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;


namespace Pars01
{
    public partial class Form1 : Form
    {
        TreeNode node = new TreeNode() { }; // узел дерева

        public Form1()
        {
            InitializeComponent();

            textBox5.Text = Properties.Settings.Default.user; 
            textBox4.Text = Properties.Settings.Default.ip;
            textBox7.Text = Properties.Settings.Default.pas;
        }

        private void button4_Click(object sender, EventArgs e) // нажали кнопку "Сохранить введенные данные"
        {
            Properties.Settings.Default.user = textBox5.Text;
            Properties.Settings.Default.ip = textBox4.Text;
            Properties.Settings.Default.pas = textBox7.Text;
            Properties.Settings.Default.Save();
        }

        private void button3_Click(object sender, EventArgs e) // нажали кнопку "Копировать фалы VipNet"
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory;

            System.Diagnostics.Process.Start(directory + "pscp.exe ", "-pw " + textBox7.Text + " " + textBox5.Text + @"@" + textBox4.Text + @":/etc/vipnet/user/iplir.conf " + directory);
            Thread.Sleep(1000);

            System.Diagnostics.Process.Start(directory + "pscp.exe ", "-pw " + textBox7.Text + " " + textBox5.Text + @"@" + textBox4.Text + @":/etc/vipnet/user/mftp.conf " + directory);
            Thread.Sleep(1000);

            Parsing2tabl(directory + "iplir.conf", "iplir.conf", richTextBox1, treeView1, dataGridView1);
            Thread.Sleep(1000);
            Parsing2tabl(directory + "mftp.conf", "mftp.conf", richTextBox2, treeView2, dataGridView2);
        }

        private void button1_Click(object sender, EventArgs e) // нажали кнопку выбора файла
        {
            openFileDialog1.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            openFileDialog1.Filter = "conf files VipNet (*.conf)|*.conf"; // фильтр выбираемых фалов
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK) // показываем диалог и если что то выбрали
            {
                try
                {
                    textBox1.Text = openFileDialog1.FileName; // пишем в текстБокс путь к файлу

                    if (openFileDialog1.SafeFileName == "iplir.conf")
                    {
                        Parsing2tabl(openFileDialog1.FileName, "iplir.conf", richTextBox1, treeView1, dataGridView1);
                        tabControl3.SelectedIndex = 0;
                    }
                    if (openFileDialog1.SafeFileName == "mftp.conf")
                    {
                        Parsing2tabl(openFileDialog1.FileName, "mftp.conf", richTextBox2, treeView2, dataGridView2);
                        tabControl3.SelectedIndex = 1;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: Не удалось прочитать файл. Вот такая вот ерунда: " + ex.Message);
                }
            }

        }

        private void подключениеToolStripMenuItem_Click(object sender, EventArgs e)// меню Подключение
        {
            Form2 f2 = new Form2();
            f2.Show();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e) // меню  "О программе"
        {
            panel1.Visible = true;
        }

        private void panel1_Click(object sender, EventArgs e) // закрываем "О программе"
        {
            panel1.Visible = false;
        }

        //------------------------------------------------------------------------------------------------------
         
        private void Parsing2tabl(string pathF, string nameF, RichTextBox RtB, TreeView TV, DataGridView DGV) // в таблицу
        {
            int sID = 0;
            string[] sP;
            double number;
            bool fl = false;

            RtB.Clear();
            TV.Nodes.Clear();
            DGV.Rows.Clear();

            StreamReader fs = new StreamReader(pathF, Encoding.GetEncoding(20866));

            while (true)
            {
                string temp = fs.ReadLine(); // Читаем строку из файла во временную переменную.

                RtB.AppendText(temp + "\n"); // пишем Оригинал
                Pars2TV(temp, TV);          // пишем дерево

                if (temp == null) break;    // Если достигнут конец файла, прерываем считывание.

                if (temp == "[id]" || temp == "[channel]")// если новая секция то добавляем новую строку в dataGridView1
                {
                    sID = DGV.Rows.Add();
                    DGV.Rows[sID].HeaderCell.Value = (sID + 1).ToString();
                }

                if (temp == "[adapter]" || temp == "[transport]")   fl = true;

                if (fl == false && temp.Intersect(@"[*]").Any() != true && temp != "" && temp.Substring(0, 1) != "#") //      temp != "[id]" && temp != "[channel]" &&
                {
                    sP = temp.Split('=');
                    if (sP[0].Trim() == "last_call" || sP[0].Trim() == "last_err")
                    {
                        Double.TryParse(sP[1].Trim(), out number);
                        DGV.Rows[sID].Cells[sP[0].Trim() + "2"].Value = ConvertFromUnixTimestamp(number);
                    }
                    if (nameF == "iplir.conf")
                        DGV.Rows[sID].Cells[sP[0].Trim() + "1"].Value += sP[1];
                    if (nameF == "mftp.conf")
                        DGV.Rows[sID].Cells[sP[0].Trim() + "2"].Value += sP[1];
                }

            }
            fs.Close();
        }
        
        private void Pars2TV(string temp , TreeView TV )                //  пишем дерево
        {
            if (temp != "" && temp != null)
            {
                if (temp.Intersect(@"[*]").Any() == true)
                {
                    node = new TreeNode() { Name = temp, Text = temp };
                    TV.Nodes.Add(node);
                }
                else
                {
                    string[] mS = temp.Split('=');

                    TreeNode Dnode = new TreeNode() { Name = mS[0], Text = temp };
                    node.Nodes.Add(Dnode);
                }
            }
        }
         
        private void toolStripButton1_Click(object sender, EventArgs e) // Экспорт в Ексель
        {
            Excel.Application excelapp = new Excel.Application();
            excelapp.SheetsInNewWorkbook = 2;
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;


            //excelapp.Workbooks.Add(Type.Missing);


            (workbook.Sheets[1] as Excel.Worksheet).Name = "_iplir.conf_";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            for (int i = 2; i < dataGridView1.RowCount + 1; i++)
            {
                for (int j = 1; j < dataGridView1.ColumnCount + 1; j++)
                {
                    worksheet.Rows[i].Columns[j] = dataGridView1.Rows[i - 2].Cells[j - 1].Value;
                }
            }

            (workbook.Sheets[2] as Excel.Worksheet).Activate(); // делаем активным второй лист

            (workbook.Sheets[2] as Excel.Worksheet).Name = "_mftp.conf_";

            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
            {
                (workbook.Sheets[2] as Excel.Worksheet).Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
            }

            for (int i = 2; i < dataGridView2.RowCount + 1; i++)
            {
                for (int j = 1; j < dataGridView2.ColumnCount + 1; j++)
                {
                    (workbook.Sheets[2] as Excel.Worksheet).Rows[i].Columns[j] = dataGridView2.Rows[i - 2].Cells[j - 1].Value;
                }
            }

            excelapp.Visible = true;

        }

        static DateTime ConvertFromUnixTimestamp(double timestamp)      // время конвертируем 
        {
            DateTime origin = new DateTime(1970, 1, 1, 7, 0, 0, 0);
            return origin.AddSeconds(timestamp);
        }

  //----  Л О Г И   --------------------------------------------------------------------------------------------------------------------------

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            openFileDialog1.Filter = "лог файлы VipNet (*.log)|*.log"; // фильтр выбираемых фалов
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK) // показываем диалог и если что то выбрали
            {
                try
                {
                    textBox2.Text = openFileDialog1.FileName; // пишем в текстБокс путь к файлу

                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Что то пошло не так: " + ex.Message);
                }
            }

        }
 

    }
}

