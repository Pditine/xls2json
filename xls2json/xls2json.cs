using System.Data;
using System.Diagnostics;
using System.Xml;
using System.Text.Json;
using System.Text.Json.Nodes;
using Excel;

namespace xls2json
{
    public partial class xls2json : Form
    {
        private const string Website = "https://github.com/Pditine/xls2json";
        private string _excelPath = "./Excel/";
        private string ExcelPath => _excelPath;
        private string FullExcelPath => Path.GetFullPath(ExcelPath);
        private string _jsonPath = "./Json/";
        private string JsonPath => _jsonPath;
        private string FullJsonPath => Path.GetFullPath(JsonPath);

        // private List<string> SelectedExcelFileNames => SelectedExcelFilePaths.Select(name => name.Split("/").Last()).ToList();
        
        private Dictionary<string,string> _excelFilePaths = new();
        
        private List<string> SelectedExcelFilePaths => CheckBoxList.CheckedItems.Cast<string>().ToList();
        
        private static Log _log;
        public static Log Log;

        public xls2json()
        {
            Log = _log ??= new Log(LogBox);
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
                Convert();
            }
            catch (Exception exception)
            {
                Log.Error(exception.Message);
                //Log.Error(exception.Message + exception.StackTrace);
            }

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void LinkLabel(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process pro = new Process();
            pro.StartInfo.UseShellExecute = true;
            pro.StartInfo.FileName = Website;
            pro.Start();
        }

        private void Init()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var excelNames = PreLoadExcelFiles();
            AddExcelFileItem(excelNames);
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

            if (!Path.Exists(JsonPath))
            {
                Directory.CreateDirectory(JsonPath);
                Log.Info($"Json文件夹不存在，已自动创建 {FullJsonPath}");
            }

            var files = Directory.GetFiles(ExcelPath);
            var fileList = new List<string>(files);
            fileList.RemoveAll(f => (!f.EndsWith(".xls") && !f.EndsWith(".xlsx"))||f.Contains("~$"));
            foreach (var file in fileList)
            {
                _excelFilePaths.Add(file.Split("/").Last(),file);
            }
            return _excelFilePaths.Keys.ToList();
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

        private void Convert()
        {
            foreach (string item in CheckBoxList.CheckedItems)
            {
                LoadExcel(_excelFilePaths[item]);
            }

            LogBoxToBottom();
        }
        
        private int FindTagIndex(string tag, DataTable table)
        {
            foreach (DataColumn column in table.Columns)
            {
                if (table.Rows[3][column].ToString().ToLower() == tag)
                {
                    return column.Ordinal;
                }
            }
            
            return -1;
        }
    }
}
