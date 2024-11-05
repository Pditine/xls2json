using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;

namespace xls2json
{
    public class Log(RichTextBox logBox)
    {
        
        public RichTextBox LogBox => logBox;

        public void Print(LogLevel level, string message)
        {
            var color = level switch
            {
                LogLevel.Info => Color.Blue,
                LogLevel.Warning => Color.Yellow,
                LogLevel.Error => Color.Red,
                LogLevel.Succese => Color.Green,
                LogLevel.Faild => Color.Red,
                _ => Color.Black
            };
            var prefix = level.ToString();
            //logBox.SelectionStart = logBox.TextLength;
            logBox.AppendText("\n[");
            logBox.SelectionColor = color;
            logBox.AppendText(prefix);
            logBox.SelectionColor = Color.Black;
            // logBox.SelectionColor = Color.Black;
            logBox.AppendText($"]{message}");
            //logBox.SelectedText = message;

            //logBox.SelectionLength = prefix.Length;


        }

        public void Info(string content)
        {
            Print(LogLevel.Info, content);
        }
        
        public void Warning(string content)
        {
            Print(LogLevel.Warning, content);
        }
        
        public void Error(string content)
        {
            Print(LogLevel.Error, content);
        }
        
        public void Succese(string content)
        {
            Print(LogLevel.Succese, content);
        }
        
        public void Faild(string content)
        {
            Print(LogLevel.Faild, content);
        }
        
        private static void OpenSpreadsheetDocumentReadonly(string filePath)
        {
            // Open a SpreadsheetDocument based on a file path.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                if (spreadsheetDocument.WorkbookPart is not null)
                {
                    // Attempt to add a new WorksheetPart.
                    // The call to AddNewPart generates an exception because the file is read-only.
                    WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

                    // The rest of the code will not be called.
                }
            }

            // Open a SpreadsheetDocument based on a stream.
            Stream stream = File.Open(filePath, FileMode.Open);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                if (spreadsheetDocument.WorkbookPart is not null)
                {
                    // Attempt to add a new WorksheetPart.
                    // The call to AddNewPart generates an exception because the file is read-only.
                    WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

                    // The rest of the code will not be called.
                }
            }

            // Open System.IO.Packaging.Package.
            Package spreadsheetPackage = Package.Open(filePath, FileMode.Open, FileAccess.Read);

            // Open a SpreadsheetDocument based on a package.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage))
            {
                if (spreadsheetDocument.WorkbookPart is not null)
                {
                    // Attempt to add a new WorksheetPart.
                    // The call to AddNewPart generates an exception because the file is read-only.
                    WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

                    // The rest of the code will not be called.
                }
            }
        }
    }


    public enum LogLevel
    {
        Info,
        Warning,
        Error,
        Succese,
        Faild
    }
}
