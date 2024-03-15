using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using ExcelDataReader;

namespace ExcelReaderTool
{
    enum FILETYPE
    {
        DEFAULT = 0,
        CONFIG = 1,
        XML = 2,
    }

    class DataStructure
    {
        public DataStructure(string remake, string name, string type)
        {
            this.Remake = remake;
            this.Name = name;
            this.Type = type;
        }
        
        public string Remake;
        public string Name;
        public string Type;
    }

    internal class ExcelReaderTool
    {
        private string _excelFolderPath = null;
        private string _configFolderPath = null;
        private string _xmlFolderPath = null;

        //key: excelName    value: time
        private Dictionary<string, string> _lastReadExcelTimeDic = new Dictionary<string, string>();

        private Queue<string> _dirtyFilesPool = new Queue<string>(8);

        //===================================================== Single =====================================================
        private ExcelReaderTool()
        {
        }

        private static ExcelReaderTool _init = null;

        public static ExcelReaderTool Init => _init ?? (_init = new ExcelReaderTool());

        //===================================================== public function =====================================================
        public void Ready()
        {
            Logger.LogPath = FindScriptRootDirectory();
            Logger.Start();

            GetConfigurationPath();
            Check_CompletionWorkPath();
        }

        public void Work()
        {
            if (Logger._errorCount > 0)
            {
                Stop();
                return;
            }

            Work_ReadExcels_WriteTargetFiles();

            Stop();
        }

        //===================================================== private function =====================================================
        private void Stop()
        {
            Logger.Stop();
        }

        private void GetConfigurationPath()
        {
            string scriptRootPath = FindScriptRootDirectory();
            string configurationPath = Directory.GetFiles(scriptRootPath, "ConfigurationInformation.txt").Length > 0
                ? Directory.GetFiles(scriptRootPath, "ConfigurationInformation.txt")[0]
                : null;
            if (configurationPath == null)
                return;

            try
            {
                using (StreamReader sr = new StreamReader(configurationPath))
                {
                    string line;

                    while ((line = sr.ReadLine()) != null)
                    {
                        ConvertConfigurationStringToPath(line);
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error("Configuration.exe Error: " + e.Message);
            }
        }

        private void Check_CompletionWorkPath()
        {
            var curDirectory = FindScriptRootDirectory();

            if (_excelFolderPath == null)
                _excelFolderPath = Directory.GetDirectories(curDirectory, "Excel").Length > 0
                    ? Directory.GetDirectories(curDirectory, "Excel")[0]
                    : null;
            if (_configFolderPath == null)
                _configFolderPath = Directory.GetDirectories(curDirectory + "/Assets", "Config").Length > 0
                    ? Directory.GetDirectories(curDirectory + "/Assets", "Config")[0]
                    : null;
            if (_xmlFolderPath == null)
                _xmlFolderPath = Directory.GetDirectories(curDirectory, "ConfigDataXML").Length > 0
                    ? Directory.GetDirectories(curDirectory, "ConfigDataXML")[0]
                    : null;

            if (_excelFolderPath == null)
            {
                try
                {
                    Directory.CreateDirectory(curDirectory + "Excel");
                }
                catch (Exception error)
                {
                    Logger.Error("Excel Reader Error :: Create read folder is error, please check! :: " + error);
                }

                _excelFolderPath = Directory.GetDirectories(curDirectory, "Excel")[0];
            }

            if (_configFolderPath == null)
            {
                try
                {
                    Directory.CreateDirectory(curDirectory + "/Assets/Config");
                }
                catch (Exception error)
                {
                    Logger.Error("Excel Reader Error :: Create write folder is error, please check! :: " + error);
                }

                _configFolderPath = Directory.GetDirectories(curDirectory + "/Assets", "Config")[0];
            }

            if (_xmlFolderPath == null)
            {
                try
                {
                    Directory.CreateDirectory(curDirectory + "/XML");
                }
                catch (Exception error)
                {
                    Logger.Error("Excel Reader Error :: Create xml folder is error, please check! :: " + error);
                }

                _xmlFolderPath = Directory.GetDirectories(curDirectory, "XML")[0];
            }
        }

        private string FindScriptRootDirectory()
        {
            var curDirectory = Directory.GetCurrentDirectory();
            string scriptRootPath = null;

            int separatorCount = 0;
            for (int i = curDirectory.Length - 1; i >= 0; i--)
            {
                if (curDirectory[i] == '\\')
                    separatorCount++;
                //Search up four levels and get the Script Root directory
                if (separatorCount == 4)
                {
                    scriptRootPath = curDirectory.Substring(0, i - 1);
                }
            }

            return scriptRootPath;
        }

        private void ConvertConfigurationStringToPath(string str)
        {
            StringBuilder tmpSB = new StringBuilder();
            string path = null;

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == '(')
                {
                    path = str.Substring(i + 1, str.Length - 2 - i);
                    break;
                }

                tmpSB.Append(str[i]);
            }

            switch (tmpSB.ToString())
            {
                case "ExcelPath":
                    _excelFolderPath = path;
                    break;
                case "ConfigPath":
                    _configFolderPath = path;
                    break;
                case "XMLPath":
                    _xmlFolderPath = path;
                    break;
                default:
                    Logger.Warning("Configuration have redundant information");
                    break;
            }
        }

        private void Work_ReadExcels_WriteTargetFiles()
        {
            var excelsPathContainer = Directory.GetFiles(_excelFolderPath, "*.xlsx");

            foreach (var excelPath in excelsPathContainer)
            {
                string excelName = GetFileNameForPath(excelPath);

                if (!_lastReadExcelTimeDic.ContainsKey(excelName))
                    _lastReadExcelTimeDic.Add(excelName, null);
#if !DEBUG
                if (_lastReadExcelTimeDic[excelName] == null ||
                    DateTime.Parse(_lastReadExcelTimeDic[excelName]) < File.GetLastWriteTime(excelPath))
                {
                    ReadExcel_WriteTargetFiles(excelPath);
                }
#endif
#if DEBUG
                if (_lastReadExcelTimeDic[excelName] == null ||
                    DateTime.Parse(_lastReadExcelTimeDic[excelName]) < DateTime.Now)
                {
                    ReadExcel_WriteTargetFiles(excelPath);
                }
#endif
            }
        }

        private string GetFileNameForPath(string path)
        {
            int leftIndex = path.Length - 1;
            bool isReachLeftBorder = false;
            int rightIndex = path.Length - 1;
            bool isReachRightBorder = false;

            while (true)
            {
                if (isReachLeftBorder && isReachRightBorder)
                    break;

                if (path[leftIndex] != '\\')
                    leftIndex--;
                else
                    isReachLeftBorder = true;

                if (path[rightIndex] != '.')
                    rightIndex--;
                else
                    isReachRightBorder = true;
            }

            return path.Substring(leftIndex + 1, rightIndex - leftIndex - 1);
        }

        private void ReadExcel_WriteTargetFiles(string excelPath)
        {
            string excelName = GetFileNameForPath(excelPath);
            _lastReadExcelTimeDic[excelName] = DateTime.Now.ToString();

            Logger.Info("Start read " + excelName);

            ClearSpecifyFile(excelName, FILETYPE.CONFIG);
            ClearSpecifyFile(excelName, FILETYPE.XML);

            _dirtyFilesPool.Enqueue(excelName);

            ReadExcel(excelPath);
        }

        private string TargetFileNameFactory(string excelName, FILETYPE type)
        {
            StringBuilder tmpSB = new StringBuilder();

            switch (type)
            {
                case FILETYPE.CONFIG:
                    return tmpSB.Append(excelName + "_Config.cs").ToString();
                case FILETYPE.XML:
                    return tmpSB.Append(excelName + "_XML.xml").ToString();
            }

            return "";
        }

        private void ClearSpecifyFile(string excelName, FILETYPE type)
        {
            string targetPath = null;

            switch (type)
            {
                case FILETYPE.CONFIG:
                    targetPath = _configFolderPath;
                    break;
                case FILETYPE.XML:
                    targetPath = _xmlFolderPath;
                    break;
            }

            var filesContainer = Directory.GetFiles(targetPath, TargetFileNameFactory(excelName, type));
            string targetFilePath = filesContainer.Length > 0 ? filesContainer[0] : null;
            if (targetFilePath == null)
                return;
            File.Delete(targetFilePath);
        }

        private void ReadExcel(string excelPath)
        {
            FileStream excelFile = File.OpenRead(excelPath);
            IExcelDataReader firstSheet = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

            //Number of valid rows on worksheet
            int rowCount = firstSheet.RowCount;
            //Number of valid columns on worksheet
            int columnCount = firstSheet.FieldCount;

            List<DataStructure> dataStructuresList = new List<DataStructure>();
            DataStructure[] dataStruturesArray = null;
            Queue<string> remakeQueue = new Queue<string>();
            Queue<string> nameQueue = new Queue<string>();
            Queue<string> typeQueue = new Queue<string>();

            for (int i = 0; i < 3; i++)
            {
                firstSheet.Read();
                for (int j = 0; j < columnCount; j++)
                {
                    if (i == 0)
                        remakeQueue.Enqueue(firstSheet.GetValue(j)?.ToString());
                    if (i == 1)
                        nameQueue.Enqueue(firstSheet.GetValue(j)?.ToString());
                    if (i == 2)
                        typeQueue.Enqueue(firstSheet.GetValue(j)?.ToString());
                }
            }

            while (remakeQueue.Count > 0)
            {
                dataStructuresList.Add(new DataStructure(remakeQueue.Dequeue(), nameQueue.Dequeue(), typeQueue.Dequeue()));
            }

            dataStruturesArray = dataStructuresList.ToArray();

            Queue<string[]> dataContainerQueue = new Queue<string[]>();
            
            for (int i = 3; i < rowCount; i++)
            {
                firstSheet.Read();
                string[] tmpStrArray = new string[columnCount];
                for (int j = 0; j < columnCount; j++)
                {
                    tmpStrArray[j] = firstSheet.GetValue(j)?.ToString();
                }
                dataContainerQueue.Enqueue(tmpStrArray);
            }
            
            BuildXML(excelPath, dataStruturesArray, dataContainerQueue);
            BuildConfig(excelPath, dataStruturesArray);
        }

        private void BuildXML(string excelPath,DataStructure[] dataStructuresArray, Queue<string[]> dataContainerQueue)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.NewLineChars = Environment.NewLine; 

            string excelName = GetFileNameForPath(excelPath);
            string xmlSavePath = _xmlFolderPath + '\\' + excelName + "_XML.xml";

            using (XmlWriter writer = XmlWriter.Create(xmlSavePath, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Root");

                while (dataContainerQueue.Count > 0)
                {
                    string[] curData = dataContainerQueue.Dequeue();

                    writer.WriteWhitespace(Environment.NewLine);
                    writer.WriteWhitespace("\t");

                    writer.WriteStartElement(dataStructuresArray[0].Name);
                    writer.WriteAttributeString("Type", dataStructuresArray[0].Type);
                    writer.WriteString(curData[0]);

                    for (int j = 1; j < dataStructuresArray.Length; j++)
                    {
                        if (dataStructuresArray[j].Name == null)
                            continue;

                        writer.WriteWhitespace(Environment.NewLine);

                        writer.WriteWhitespace("\t\t");
                        writer.WriteStartElement(dataStructuresArray[j].Name);
                        writer.WriteAttributeString("Type", dataStructuresArray[j].Type);
                        writer.WriteString(curData[j] ?? "Null");

                        writer.WriteEndElement();
                        if (j == dataStructuresArray.Length - 1)
                        {
                            writer.WriteWhitespace(Environment.NewLine);
                            writer.WriteWhitespace("\t");
                        }
                    }
                    
                    writer.WriteEndElement();
                    writer.WriteWhitespace(Environment.NewLine);
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void BuildConfig(string excelPath,DataStructure[] dataStructuresArray)
        {
            
        }
    }
}