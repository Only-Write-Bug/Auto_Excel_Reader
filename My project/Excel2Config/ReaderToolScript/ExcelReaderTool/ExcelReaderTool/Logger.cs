using System;
using System.IO;
using System.Text;

namespace ExcelReaderTool
{
    internal static class Logger
    {
        private static string _logPath = null;
        private static string _logsFolderPath = null;
        
        public static int _errorCount = 0;

        private static StreamWriter _curWriter = null;
        
        public static string LogPath
        {
            get => _logPath;
            set
            {
                if (_logPath == null)
                    _logPath = value;
            }
        }

        //===================================================== public function =====================================================
        public static void Start()
        {
            if (_logPath == null)
                return;
            if(_logsFolderPath == null)
                if (!CreateLogsFolder())
                    return;

            _errorCount = 0;
            CreateNotePadFile();
        }

        public static int Stop()
        {
            _curWriter.Close();
            return _errorCount;
        }

        public static void Info(string str)
        {
            _curWriter.WriteLine("[Info] :: " + str);
        }
        
        public static void Warning(string str)
        {
            _curWriter.WriteLine("[Warning] :: " + str);
        }
        
        public static void Error(string str)
        {
            _curWriter.WriteLine("[Error] :: " + str);
            _errorCount++;
            Console.WriteLine("Error :: " + str);
        }

        //===================================================== private function =====================================================
        private static bool CreateLogsFolder()
        {
            try
            {
                Directory.CreateDirectory(_logPath + "\\ExcelReaderTool_WorkLogs");
            }
            catch (Exception error)
            {
                Console.WriteLine("Logger :: Error :: Create logs folder is error, please check! :: " + error);
                return false;
            }

            _logsFolderPath = Directory.GetDirectories(_logPath, "ExcelReaderTool_WorkLogs")?[0];
            return true;
        }
        
        private static void CreateNotePadFile()
        {
            _curWriter = new StreamWriter(_logsFolderPath + "\\" + CreateLogFileName() + ".txt");
        }
        
        private static string CreateLogFileName()
        {
            StringBuilder tmpSB = new StringBuilder();
            string nowTime = DateTime.Now.ToString();

            foreach (var character in nowTime)
            {
                if (character == '/' || character == ' ' || character == ':')
                {
                    tmpSB.Append('_');
                    continue;
                }

                tmpSB.Append(character);
            }
            
            return tmpSB.ToString();
        }
    }
}