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
            fileList.RemoveAll(f => !f.EndsWith(".xls") && !f.EndsWith(".xlsx"));
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

        private void LoadExcel(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                Log.Error($"文件不存在 {excelPath}");
                return;
            }
            Log.Info($"加载文件 {excelPath}");
            using var fs = File.Open(excelPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelDataReader;
            if(excelPath.EndsWith(".xls"))
            {
                excelDataReader = ExcelReaderFactory.CreateBinaryReader(fs);
            }
            else
            {
                excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            }
            var result = excelDataReader.AsDataSet();
            if(result == null)
            {
                Log.Error($"文件读取失败 {excelPath} : {excelDataReader.ExceptionMessage}");
                return;
            }
            
            foreach (DataTable table in result.Tables)
            {
                LoadSheet(excelPath,table);
            }
            
            Log.Succese("文件处理完成");
            fs.Close();
        }
        
        private void LoadSheet(string filePath, DataTable table)
        {
            var jsonObject = new JsonObject();
            var options = new JsonSerializerOptions { WriteIndented = true };
            Log.Info("处理表格:" + table.TableName);
            jsonObject["Meta"] = GetFileHead(filePath, table.TableName);
            var jsonString = jsonObject.ToJsonString(options);

            int keyIndex = -1;

            // DataRow tagRow = table.Rows[0];
            // foreach (DataColumn column in table.Columns)
            // {
            //     Log.Debug(tagRow[column].ToString());
            //     if (tagRow[column].ToString().ToLower() == "key")
            //     {
            //         keyIndex = column.Ordinal;
            //         break;
            //     }
            // }
            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn col in table.Columns)
                {
                    Log.Debug(row[col].ToString());   
                }
            }
            
            if(keyIndex == -1)
            {
                Log.Error("未找到主键key");
                return;
            }
            Log.Debug(keyIndex.ToString());
            var jsonPath = JsonPath + table.TableName + ".json";
            File.Create(jsonPath).Close();
            File.WriteAllText(jsonPath, jsonString);
            
            Log.Succese($"已生成文件 {jsonPath}");
        }

        private JsonNode GetFileHead(string filePath,string tableName)
        {
            var meta = new JsonObject();
            meta["Description"] = "This file is generated by xls2json, do not modify it manually.";
            meta["CopyRight"] = "Pditine";
            meta["Official Website"] = Website;
            meta["File"] = filePath;
            meta["Sheet"] = tableName;
            return meta;
        }

        private void Convert()
        {
            foreach (string item in CheckBoxList.CheckedItems)
            {
                LoadExcel(_excelFilePaths[item]);
            }

            LogBoxToBottom();
        }
    }
}
