using System;
using System.Diagnostics;
using System.Security.Policy;

namespace xls2json
{
    public partial class xls2json : Form
    {
        private const string ExcelPath = "Excel";
        private Log _log;
        public Log Log
        {
            get
            {
                if (_log == null)
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
            Init();
        }

        private void Convert_Click(object sender, EventArgs e)
        {
            Log.Info("转换成功");
            LogBoxToBottom();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void LinkLabel(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process pro = new Process();
            pro.StartInfo.UseShellExecute = true;
            pro.StartInfo.FileName = "https://github.com/Pditine/xls2json";
            pro.Start();
        }

        private void Init()
        {
            var files = PreLoadExcelFiles();
            AddExcelFileItem(files);
        }

        private List<string> PreLoadExcelFiles()
        {
            if (!Directory.Exists(ExcelPath))
                Directory.CreateDirectory(ExcelPath);
            var files = Directory.GetFiles(ExcelPath);
            var fileList = new List<string>(files);
            fileList.RemoveAll(f => !f.EndsWith(".xls") && !f.EndsWith(".xlsx"));
            return fileList;
        }

        private void AddExcelFileItem(List<string> files)
        {
            foreach (var file in files)
            {
                AddExcelFileItem(file);
            }
        }

        private void AddExcelFileItem(string file)
        {
            CheckBoxList.Items.Add(file);
        }

        private void LogBoxToBottom()
        {
            LogBox.SelectionStart = LogBox.Text.Length;
            LogBox.SelectionLength = 0;
            LogBox.ScrollToCaret();
        }
    }
}
