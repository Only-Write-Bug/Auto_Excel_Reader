using System;
using System.IO;
using System.Text;

namespace ExcelReaderTool
{
    internal static class Logger
    {
        private static string logPath = null;
        private static string logsFloderPath = null;

        private static StreamWriter curWriter = null;
        
        public static string LogPath
        {
            get { return logPath; }
            set
            {
                if (logPath == null)
                    logPath = value;
            }
        }

        //===================================================== public function =====================================================
        public static void Start()
        {
            if (logPath == null)
                return;
            if(logsFloderPath == null)
                if (!CreateLogsFloder())
                    return;

            CreateNotePadFile();
        }

        public static void Stop()
        {
            curWriter.Close();
        }

        public static void Info(string str)
        {
            curWriter.WriteLine("[Info] :: " + str);
        }
        
        public static void Warning(string str)
        {
            curWriter.WriteLine("[Warning] :: " + str);
        }
        
        public static void Error(string str)
        {
            curWriter.WriteLine("[Error] :: " + str);
        }

        //===================================================== private function =====================================================
        private static bool CreateLogsFloder()
        {
            try
            {
                Directory.CreateDirectory(logPath + "\\ExcelReaderTool_WorkLogs");
            }
            catch (Exception error)
            {
                Console.WriteLine("Logger :: Error :: Create logs floder is error, please check! :: " + error);
                return false;
            }

            logsFloderPath = Directory.GetDirectories(logPath, "ExcelReaderTool_WorkLogs")?[0];
            return true;
        }
        
        private static void CreateNotePadFile()
        {
            curWriter = new StreamWriter(logsFloderPath + "\\" + CreateLogFileName() + ".txt");
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