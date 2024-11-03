using System;
using System.Diagnostics;
using System.Security.Policy;

namespace xls2json
{
    public partial class xls2json : Form
    {
        private Log _log;
        public Log Log
        {
            get
            {
                if( _log == null )
                    _log = new Log(LogBox);
                return _log;
            }
        }

        public xls2json()
        {
            InitializeComponent();
        }

        private void Xls2json_Load(object sender, EventArgs e)
        {
            Log.Succese("Application started");
        }

        private void Convert_Click(object sender, EventArgs e)
        {
            Log.Info("转换成功");
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


    }


}
