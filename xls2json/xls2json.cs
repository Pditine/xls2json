using System;
using System.Data;
using System.Diagnostics;
using System.Xml;
using Excel;

namespace xls2json
{
    public partial class xls2json : Form
    {
        private List<string> _excelNames;
        private string _excelPath = "./Excel/";
        private string ExcelPath => _excelPath;
        private string FullExcelPath => Path.GetFullPath(ExcelPath);
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
            try
            {
                ConvertExcel();
            }
            catch (Exception exception)
            {
                Log.Error(exception.Message + exception.StackTrace);
                throw;
            }

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
            _excelNames = PreLoadExcelFiles();
            AddExcelFileItem(_excelNames);
        }

        private List<string> PreLoadExcelFiles()
        {
            if (!Path.Exists(ExcelPath))
            {
                Directory.CreateDirectory(ExcelPath);
                Log.Info($"Excel文件夹不存在，已自动创建 {FullExcelPath}");
            }else
            {
                Log.Info($"从 {FullExcelPath} 加载文件");
            }
            
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

        private void LoadExcel(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                Log.Error($"文件不存在 {excelPath}");
                return;
            }
            using var fs = File.Open(excelPath, FileMode.Open, FileAccess.Read);
            var excelDataReader = ExcelReaderFactory.CreateBinaryReader(fs);
            //todo: 分支判断
            //var excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            var result = excelDataReader.AsDataSet();
            if(result == null)
            {
                Log.Error($"文件读取失败 {excelPath}");
                return;
            }
            
            foreach (DataTable table in result.Tables)
            {
                LoadSheet(table);
            }
            
            fs.Close();
        }
        
        private void LoadSheet(DataTable table)
        {
            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn column in table.Columns)
                {
                    var value = row[column];
                    Log.Debug(value.ToString());
                }
            }
        }

        private void WriteInJson()
        {
            
        }

        private void ConvertExcel()
        {
            foreach (string excelName in _excelNames)
            {
                LoadExcel(excelName);
            }

            LogBoxToBottom();
        }
    }
}
