using System;
using System.Diagnostics;
using System.Security.Policy;

namespace xls2json
{
    public partial class xls2json : Form
    {
        public xls2json()
        {
            InitializeComponent();
        }

        private void Xls2json_Load(object sender, EventArgs e)
        {
            Console.WriteLine("123");
            Log(LogLevel.Info,"123");
        }

        private void Convert_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process pro = new Process();
            pro.StartInfo.UseShellExecute = true;
            pro.StartInfo.FileName = "https://purpleditine.top/";
            pro.Start();
        }

        private void Log(LogLevel level, string content)
        {
            LogBox.AppendText($"\n[<color=00ff00>{level}</color>]");
        }
    }


    enum LogLevel
    {
        Info,
        Warning,
        Error,
        Succese,
        Faild
    }
}
