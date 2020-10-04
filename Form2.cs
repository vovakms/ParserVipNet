using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pars01
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory;
            textBox6.Text = directory;
            textBox2.Text = @"fssf11@10.27.11.129:/etc/vipnet/user/iplir.conf " + directory ;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory;
            string str2 = textBox2.Text ;
            // textBox2.Text = "\"" + directory + "pscp.exe" + "\" " + str2;

            System.Diagnostics.Process.Start(directory + "pscp.exe ", @"fssf11@10.27.11.129:/etc/vipnet/user/iplir.conf " + directory);

            System.Diagnostics.Process.Start(directory + "pscp.exe ", @"fssf11@10.27.11.129:/etc/vipnet/user/mftp.conf " + directory);
        }




    }
}
