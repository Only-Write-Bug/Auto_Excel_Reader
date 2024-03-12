using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using System.Data;
using System;
using ExcelDataReader;
using System.Text;
using Config;
using System.Reflection;
using System.Xml;

namespace Config
{
    public static class TypeConversionTool
    {
        public static dynamic XMLData2ConfigData(string type, string data)
        {
            switch (type)
            {
                case "int":
                    return TypeConversionTool.Str2Int(data);
                case "short":
                    return TypeConversionTool.Str2Short(data);
                case "long":
                    return TypeConversionTool.Str2Long(data);
                case "float":
                    return TypeConversionTool.Str2float(data);
                case "double":
                    return TypeConversionTool.Str2Double(data);
                case "bool":
                    return TypeConversionTool.Str2Bool(data);
                case "char":
                    return TypeConversionTool.Str2Char(data);
                case "Vector2":
                    return TypeConversionTool.Str2Vector2(data);
                case "Vector3":
                    return TypeConversionTool.Str2Vector3(data);
                case "int[]":
                    return TypeConversionTool.Str2BaseTypeArray<int>(data);
                case "short[]":
                    return TypeConversionTool.Str2BaseTypeArray<short>(data);
                case "long[]":
                    return TypeConversionTool.Str2BaseTypeArray<long>(data);
                case "float[]":
                    return TypeConversionTool.Str2BaseTypeArray<float>(data);
                case "double[]":
                    return TypeConversionTool.Str2BaseTypeArray<double>(data);
                case "bool[]":
                    return TypeConversionTool.Str2BaseTypeArray<bool>(data);
                case "char[]":
                    return TypeConversionTool.Str2BaseTypeArray<char>(data);
                case "string[]":
                    return TypeConversionTool.Str2BaseTypeArray<string>(data);
                case "Vector2[]":
                    return TypeConversionTool.Str2Vector2Array(data);
                case "Vector3[]":
                    return TypeConversionTool.Str2Vector3Array(data);
                default:
                    break;
            }

            return null;
        }

        public static int Str2Int(string str)
        {
            int.TryParse(str, out int result);
            return result;
        }

        public static short Str2Short(string str)
        {
            short.TryParse(str, out short result);
            return result;
        }

        public static long Str2Long(string str)
        {
            long.TryParse(str, out long result);
            return result;
        }

        public static float Str2float(string str)
        {
            float.TryParse(str, out float result);
            return result;
        }

        public static double Str2Double(string str)
        {
            double.TryParse(str, out double result);
            return result;
        }

        public static bool Str2Bool(string str)
        {
            bool.TryParse(str, out bool result);
            return result;
        }

        public static char Str2Char(string str)
        {
            char.TryParse(str, out char result);
            return result;
        }

        public static Vector2 Str2Vector2(string str)
        {
            Vector2 result = new Vector2();

            if (str.Length <= 0 || str == null || str.ToUpper() == "NULL")
                return result;

            StringBuilder tmpSB = new StringBuilder();
            for (int i = 1; i < str.Length - 1; i++)
            {
                tmpSB.Append(str[i]);
            }

            var numArray = Str2BaseTypeArray<float>(tmpSB.ToString());
            result.x = float.Parse(numArray.GetValue(0).ToString());
            result.y = float.Parse(numArray.GetValue(1).ToString());

            return result;
        }

        public static Vector3 Str2Vector3(string str)
        {
            Vector3 result = new Vector3();

            if (str.Length <= 0 || str == null || str.ToUpper() == "NULL")
                return result;

            StringBuilder tmpSB = new StringBuilder();
            for (int i = 1; i < str.Length - 1; i++)
            {
                tmpSB.Append(str[i]);
            }

            var numArray = Str2BaseTypeArray<float>(tmpSB.ToString());
            result.x = float.Parse(numArray.GetValue(0).ToString());
            result.y = float.Parse(numArray.GetValue(1).ToString());
            result.z = float.Parse(numArray.GetValue(2).ToString());

            return result;
        }

        public static Vector2[] Str2Vector2Array(string str)
        {
            List<Vector2> tmpList = new List<Vector2>();
            StringBuilder tmpSB = new StringBuilder();
            bool vectorIsOver = false;

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == ',' && vectorIsOver)
                {
                    tmpList.Add(Str2Vector2(tmpSB.ToString()));
                    vectorIsOver = false;
                    tmpSB.Clear();
                    continue;
                }

                tmpSB.Append(str[i]);

                if (str[i] == ']')
                    vectorIsOver = true;

                if (i == str.Length - 1)
                {
                    tmpList.Add(Str2Vector2(tmpSB.ToString()));
                    break;
                }
            }

            return tmpList.ToArray();
        }

        public static Vector3[] Str2Vector3Array(string str)
        {
            List<Vector3> tmpList = new List<Vector3>();
            StringBuilder tmpSB = new StringBuilder();
            bool vectorIsOver = false;

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == ',' && vectorIsOver)
                {
                    tmpList.Add(Str2Vector3(tmpSB.ToString()));
                    vectorIsOver = false;
                    tmpSB.Clear();
                    continue;
                }

                tmpSB.Append(str[i]);

                if (str[i] == ']')
                    vectorIsOver = true;

                if (i == str.Length - 1)
                {
                    tmpList.Add(Str2Vector3(tmpSB.ToString()));
                    break;
                }
            }

            return tmpList.ToArray();
        }

        //apply to int、float、short、long、double、bool、char、string
        public static Array Str2BaseTypeArray<T>(string str)
        {
            if (str.Length <= 0)
                return null;

            Queue<string> dataQueue = new Queue<string>(8);
            StringBuilder tmpSB = new StringBuilder();

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == ',')
                {
                    dataQueue.Enqueue(tmpSB.ToString());
                    tmpSB.Clear();
                    continue;
                }

                tmpSB.Append(str[i]);
            }
            dataQueue.Enqueue(tmpSB.ToString());

            var dataArray = Array.CreateInstance(typeof(T), dataQueue.Count);

            for (int i = 0; i < dataArray.Length; i++)
            {
                dataArray.SetValue((T)Convert.ChangeType(dataQueue.Dequeue(), typeof(T)), i);
            }

            return dataArray;
        }
    }
}

public class ExcelReaderTool
{
    private static string _readFloderPath = null;
    private static string _writeFloderPath = null;
    private static string _xmlFloderPath = null;
    public static string get_xmlFloderPath
    {
        get
        {
            return _xmlFloderPath;
        }
    }

    private static string _configName = null;
    private static string _configScripName = null;
    private static StringBuilder _dataClass = new StringBuilder();

    private static string[] _nameContainer = null;
    private static string[] _typeContainer = null;

    [MenuItem("Tools/Excel Reader/Read Excel")]
    public static void ExcelReader()
    {
        if (!InitWorkFloderStream())
        {
            Debug.LogError("Excel Reader Error :: Work Floder Stream is Error, stop work!!!");
            return;
        }
        Work();
    }

    [MenuItem("Tools/Excel Reader/Write Config")]
    public static void WriteConfig()
    {
        if(_writeFloderPath == null)
        {
            var curDirectory = Directory.GetCurrentDirectory();
            _writeFloderPath = Directory.GetDirectories(curDirectory + "/Assets", "Config").Length > 0 ? Directory.GetDirectories(curDirectory + "/Assets", "Config")[0] : null;
        }

        foreach(var configFile in Directory.GetFiles(_writeFloderPath, "*.cs"))
        {
            StringBuilder tmpSB = new StringBuilder();
            GetFileName(configFile, tmpSB);
            Debug.Log(tmpSB.ToString());
        }
    }

    private static bool InitWorkFloderStream()
    {
        //Get project path
        var curDirectory = Directory.GetCurrentDirectory();

        if (_readFloderPath == null)
            _readFloderPath = Directory.GetDirectories(curDirectory, "Excel").Length > 0 ? Directory.GetDirectories(curDirectory, "Excel")[0] : null;
        if (_writeFloderPath == null)
            _writeFloderPath = Directory.GetDirectories(curDirectory + "/Assets", "Config").Length > 0 ? Directory.GetDirectories(curDirectory + "/Assets", "Config")[0] : null;
        if (_xmlFloderPath == null)
            _xmlFloderPath = Directory.GetDirectories(curDirectory, "ConfigDataXML").Length > 0 ? Directory.GetDirectories(curDirectory, "ConfigDataXML")[0] : null;

        if (_readFloderPath == null)
        {
            try
            {
                Directory.CreateDirectory(curDirectory + "Excel");
            }
            catch (Exception error)
            {
                Debug.LogError("Excel Reader Error :: Create read floder is error, please check! :: " + error);
                return false;
            }
            _readFloderPath = Directory.GetDirectories(curDirectory, "Excel")[0];
        }
        if (_writeFloderPath == null)
        {
            try
            {
                Directory.CreateDirectory(curDirectory + "/Assets/Config");
            }
            catch (Exception error)
            {
                Debug.LogError("Excel Reader Error :: Create write floder is error, please check! :: " + error);
                return false;
            }
            _writeFloderPath = Directory.GetDirectories(curDirectory + "/Assets", "Config")[0];
        }
        if (_xmlFloderPath == null)
        {
            try
            {
                Directory.CreateDirectory(curDirectory + "/ConfigDataXML");
            }
            catch (Exception error)
            {
                Debug.LogError("Excel Reader Error :: Create xml floder is error, please check! :: " + error);
                return false;
            }
            _readFloderPath = Directory.GetDirectories(curDirectory, "ConfigDataXML")[0];
        }

        return true;
    }

    private static void Work()
    {
        //Get all excel files
        var excelFilesNameContainer = Directory.GetFiles(_readFloderPath, "*.xlsx");

        ClearConfig();
        ClearXML();

        foreach (var fileStream in excelFilesNameContainer)
        {
            string excelFilePath = fileStream;
            FileStream excelFile = File.OpenRead(excelFilePath);
            IExcelDataReader firstSheet = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

            if (firstSheet == null)
            {
                Debug.LogError($"{fileStream} convert to Excel Sheet fail��");
                continue;
            }

            FileStream csFile = new FileStream(FileNameFactory(fileStream), FileMode.Create);

            ExcelReaderTool.ReadExcelSheet_WriteCSFile(firstSheet, csFile);
        }
        Debug.Log("Excel Reader Work is over!");
    }

    private static void ClearConfig()
    {
        //Get all config files
        var configFilesNameContainer = Directory.GetFiles(_writeFloderPath, "*.*");

        foreach (var file in configFilesNameContainer)
        {
            File.Delete(file);
        }
    }

    private static void ClearXML()
    {
        //Get all config files
        var configFilesNameContainer = Directory.GetFiles(_xmlFloderPath, "*.*");

        foreach (var file in configFilesNameContainer)
        {
            File.Delete(file);
        }
    }

    private static string FileNameFactory(string stream)
    {
        StringBuilder name = new StringBuilder();

        GetFileName(stream, name);
        _configName = name.ToString();
        name.Append("Config.cs");

        return _writeFloderPath + '\\' + name.ToString();
    }

    private static void GetFileName(string stream, StringBuilder stringBuilder)
    {
        char[] charArray = stream.ToCharArray();
        foreach (char c in charArray)
        {
            if (c == '\\')
            {
                stringBuilder.Clear();
                continue;
            }

            if (c == '.')
            {
                break;
            }

            stringBuilder.Append(c);
        }
    }

    private static void ReadExcelSheet_WriteCSFile(IExcelDataReader Sheet, FileStream file)
    {
        //Number of valid rows on worksheet
        int rowCount = Sheet.RowCount;
        //Number of valid columns on worksheet
        int columnCount = Sheet.FieldCount;

        _nameContainer = new string[columnCount];
        _typeContainer = new string[columnCount];

        //befor WriteDataClass(), Init Data
        string[][] dataInfo = new string[3][];
        for (int i = 0; i < 3; i++)
        {
            dataInfo[i] = new string[columnCount];
            Sheet.Read();
            for (int j = 0; j < columnCount; j++)
            {
                dataInfo[i][j] = Sheet.GetValue(j)?.ToString();
                if (i == 1)
                    _nameContainer[j] = dataInfo[i][j];
                if (i == 2)
                    _typeContainer[j] = dataInfo[i][j];
            }
        }

        WriteDataClass(dataInfo);

        InitConfigCS(file);

        //Data Storage, Write DataXML File
        XmlDocument xml = new XmlDocument();
        XmlElement rootElement = xml.CreateElement("Root");
        xml.AppendChild(rootElement);
        for (int i = 3; i < rowCount; i++)
        {
            Sheet.Read();

            XmlElement sonElement = xml.CreateElement(_nameContainer[0]);
            sonElement.SetAttribute("Type", _typeContainer[0]);
            sonElement.InnerText = Sheet.GetValue(0).ToString();

            for (int j = 1; j < columnCount; j++)
            {
                var v = Sheet.GetValue(j);

                if (_nameContainer?[j] == null || _typeContainer?[j] == null)
                    continue;

                XmlElement grandsonElement = xml.CreateElement(_nameContainer[j]);
                grandsonElement.SetAttribute("Type", _typeContainer[j]);
                grandsonElement.InnerText = v == null ? "NULL" : v.ToString();
                sonElement.AppendChild(grandsonElement);
            }
            rootElement.AppendChild(sonElement);
        }
        xml.Save(_xmlFloderPath + "/" + _configScripName + "_XML.xml");

        file.Close();

        //Begin Config init, cause if start init on useing, maybe too long time
        //var configScriptType = Type.GetType("Config." + _configScripName);
        //if (configScriptType == null)
        //{
        //    Debug.LogError($"Excel Reader Error :: doesn't find name is Config.{_configScripName}'s script");
        //}
        //else
        //{
        //    PropertyInfo propertyInfo = configScriptType.GetProperty("Init");
        //    MethodInfo methodInfo = propertyInfo.GetGetMethod();
        //    object configInstance = Activator.CreateInstance(configScriptType);
        //    methodInfo.Invoke(configInstance, null);
        //}

        _configScripName = null;
    }

    private static void InitConfigCS(FileStream file)
    {
        WriteReferencePackage(file);
        WriteStaticConfig(file);
    }

    private static void WriteReferencePackage(FileStream file)
    {
        String data = "using UnityEngine;\n" +
            "using System;\n" +
            "using System.Xml;\n" +
            "using System.IO;\n" +
            "using System.Collections.Generic;\n\n";

        byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(data);
        file.Write(byteArray, 0, byteArray.Length);
    }

    private static void WriteStaticConfig(FileStream file)
    {
        StringBuilder nameSB = new StringBuilder();
        GetFileName(file.Name, nameSB);
        _configScripName = nameSB.ToString();

        string data =
              "namespace Config\n" +
              "{\n" +
              $"     public class {_configScripName}\n" +
              "     {\n" +
              "         public " + _configScripName + "() \n" +
              "         {\n" +
              $"             this._dataContainer = new Dictionary<int, {_configName}Data>(8);\n" +
              "         }\n\n" +
              $"         private static {_configScripName} _init = null;\n" +
              $"         public static {_configScripName} Init\n" +
              "         {\n" +
              "             get\n" +
              "             {\n" +
              "                 if (_init == null)\n" +
              "                 {\n" +
              $"                    _init = new {_configScripName}();\n" +
              $"                    _init.OnCreate();\n" +
              "                 }\n" +
              "                 return _init;\n" +
              "             }\n" +
              "         }\n\n" +
              "         //key is Data's ID\n" +
              $"        private Dictionary<int, {_configName}Data> _dataContainer = null;\n\n" +
              "         private void OnCreate()\n" +
              "         {\n" +
              $"             string xmlFileName = \"{_configScripName}_XML.xml\";\n" +
              "             string filePath = Path.Combine(ExcelReaderTool.get_xmlFloderPath, xmlFileName);\n" +
              "             if(File.Exists(filePath))\n" +
              "             {\n" +
              "                 XmlDocument xml = new XmlDocument();\n" +
              "                 xml.Load(filePath);\n\n" +
              "                 XmlNode root = xml.DocumentElement;\n\n" +
              "                 foreach(XmlNode sonNode in root.ChildNodes)\n" +
              "                 {\n" +
              "                     Debug.Log(sonNode.Name);\n" +
              "                     foreach(XmlNode grandsonNode in sonNode.ChildNodes)\n" +
              "                     {\n" +
              "                         if(grandsonNode.Name == \"#text\")\n" +
              "                             continue;\n" +
              "                         Debug.Log(grandsonNode.Name);\n" +
              "                     }\n" +
              "                 }\n" +
              "             }\n" +
              "             else\n" +
              "             {\n" +
              $"                 Debug.LogError(\"Excel Reader Error :: {_configScripName}_XML.xml is not find;\");\n" +
              "             }\n" +
              "         }\n\n" +
              "" +
              "     }\n" +
              _dataClass.ToString() +
              "\n" +
              "}\n";

        byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(data);
        file.Write(byteArray, 0, byteArray.Length);

        _dataClass.Clear();
        _configName = null;
    }

    private static void WriteDataClass(string[][] dataInfo)
    {
        _dataClass.Append(
              $"\n     public class {_configName}Data\n" +
              "     {\n");

        int infoWidth = dataInfo[0].Length;

        //write init method
        StringBuilder initMethodArgsString = new StringBuilder();
        StringBuilder initMembersString = new StringBuilder();

        initMethodArgsString.Append("#pragma warning disable CS0414\n");
        for (int i = 0; i < infoWidth; i++)
        {
            if (dataInfo[1][i] == null && dataInfo[2][i] == null)
            {
                continue;
            }

            if (dataInfo[1][i] == null || dataInfo[2][i] == null)
            {
                Debug.LogError($"Excel Reader Error :: {_configName} :: Excel data incomplete information, please check your excel stats name or type is incomplete?");
                _dataClass.Append(
                "     }\n");
                return;
            }

            //TODO: Check whether the type is valid

            initMethodArgsString.Append($"         private {dataInfo[2][i]} _default_{dataInfo[1][i]} = {GetTypeNormalValue(dataInfo[2][i])};\n");
        }
        initMethodArgsString.Append("#pragma warning restore CS0414\n\n");

        _dataClass.Append(initMethodArgsString.ToString() +
             $"         public {_configName}Data()\n" +
              "         {\n" +
              "             \n" +
              initMembersString.ToString() +
              "         }\n\n");

        //write member
        for (int i = 0; i < infoWidth; i++)
        {
            if (dataInfo[0][i] != null)
            {
                _dataClass.Append(
                    $"          //{dataInfo[0][i]}\n");
            }

            if (dataInfo[1][i] == null && dataInfo[2][i] == null)
            {
                continue;
            }

            _dataClass.Append(
                $"          private {dataInfo[2][i]} _{dataInfo[1][i]};\n" +
                $"          public {dataInfo[2][i]} {dataInfo[1][i]}\n" +
                 "          {\n" +
                 "              get { return " +
                $"this._{dataInfo[1][i]}; " +
                 "}\n" +
                 "          }\n\n");
        }

        _dataClass.Append(
              "     }\n");

    }

    private static string GetTypeNormalValue(string type)
    {
        switch (type)
        {
            case "short":
            case "long":
            case "int":
            case "float":
            case "double":
                return "0";
            case "bool":
                Debug.LogWarning($"Excel Reader Error :: {_configName} :: The file has an unset type of bool, please check, the default setting is false!");
                return "false";
            case "char":
                return "''";
            case "Vector2":
                return "new Vector2()";
            case "Vector3":
                return "new Vector3()";
            default:
                Debug.LogWarning("Excel Reader Tool Warning :: Data Type is not definition, please check!");
                break;
        }

        return "null";
    }

}