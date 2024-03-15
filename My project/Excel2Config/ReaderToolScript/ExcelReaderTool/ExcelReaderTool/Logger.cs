using System;
using System.IO;
using System.Text;

namespace ExcelReaderTool
{
    internal static class Logger
    {
        private static string _logPath = null;
        private static string _logsFloderPath = null;

        private static StreamWriter _curWriter = null;
        
        public static string LogPath
        {
            get { return _logPath; }
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
            if(_logsFloderPath == null)
                if (!CreateLogsFloder())
                    return;

            CreateNotePadFile();
        }

        public static void Stop()
        {
            _curWriter.Close();
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
            Console.WriteLine("Error :: " + str);
        }

        //===================================================== private function =====================================================
        private static bool CreateLogsFloder()
        {
            try
            {
                Directory.CreateDirectory(_logPath + "\\ExcelReaderTool_WorkLogs");
            }
            catch (Exception error)
            {
                Console.WriteLine("Logger :: Error :: Create logs floder is error, please check! :: " + error);
                return false;
            }

            _logsFloderPath = Directory.GetDirectories(_logPath, "ExcelReaderTool_WorkLogs")?[0];
            return true;
        }
        
        private static void CreateNotePadFile()
        {
            _curWriter = new StreamWriter(_logsFloderPath + "\\" + CreateLogFileName() + ".txt");
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