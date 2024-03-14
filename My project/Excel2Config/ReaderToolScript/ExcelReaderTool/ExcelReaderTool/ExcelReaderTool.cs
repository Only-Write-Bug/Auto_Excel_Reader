using System;
using System.IO;

namespace ExcelReaderTool
{
    internal class ExcelReaderTool
    {
        private string readPath = null;
        private string configPath = null;
        private string xmlPath = null;
        private string logPath = null;
        
        private int warningCount = 0;
        public int Get_WarningCount
        {
            get { return warningCount; }
        }
        
        private int errorCount = 0;
        public int Get_ErrorCount
        {
            get { return errorCount; }
        }

        //===================================================== Singel =====================================================
        private ExcelReaderTool() 
        {
            Ready();
        }

        private static ExcelReaderTool _init = null;

        public static ExcelReaderTool Init
        {
            get
            {
                if(_init == null)
                {
                    _init = new ExcelReaderTool();
                }
                return _init;
            }
        }

        //===================================================== public function =====================================================
        public void CheckWorkPath()
        {
            if (readPath != null)
                Console.WriteLine(readPath);
            else
            {
                Logger.Error("Error :: readPath is null");
                errorCount++;
            }
            if (configPath != null)
                Console.WriteLine(configPath);
            else
            {
                Logger.Error("Error :: configPath is null");
                errorCount++;
            }
            if (xmlPath != null)
                Console.WriteLine(xmlPath);
            else
            {
                Logger.Error("Error :: xmlPath is null");
                errorCount++;
            }
        }

        //===================================================== private function =====================================================
        private void Ready()
        {
            if(!GetConfigurationPath())
            {
                
            }

            Logger.LogPath = FindScriptRootDirectory();
            Logger.Start();
        }

        private bool GetConfigurationPath()
        {
            string scriptRootPath = FindScriptRootDirectory();
            string configurationPath = Directory.GetFiles(scriptRootPath, "ConfigurationInformation.txt").Length > 0 ? 
                Directory.GetFiles(scriptRootPath, "ConfigurationInformation.txt")[0] : 
                null;
            if (configurationPath == null)
                return false;

            try
            {
                using (StreamReader sr = new StreamReader(configurationPath))
                {
                    string line;

                    while ((line = sr.ReadLine()) != null)
                    {
                        Console.WriteLine(line);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("读取文件时发生错误: " + e.Message);
            }

            return true;
        }

        private void CreateWorkPath()
        {

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
    }
}