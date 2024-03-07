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

namespace Config { }

public class ExcelReaderTool
{
    private static string _readFloderPath = null;
    private static string _writeFloderPath = null;
    private static string _xmlFloderPath = null;

    private static string _configName = null;
    private static string _configScripName = null;
    private static StringBuilder _dataClass = new StringBuilder();

    private static string[] _typeContainer = null;

    [MenuItem("Tools/Excel Reader")]
    public static void ExcelReader()
    {
        if(!InitWorkFloderStream())
        {
            Debug.LogError("Excel Reader Error :: Work Floder Stream is Error, stop work!!!");
            return;
        }
        Work();
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
                if (i == 2)
                {
                    _typeContainer[j] = dataInfo[i][j];
                }
            }
        }

        WriteDataClass(dataInfo);

        InitConfigCS(file);

        //Data Storage
        XmlDocument xml = new XmlDocument();
        for (int i = 3; i < rowCount; i++)
        {
            Sheet.Read();
            for (int j = 0; j < columnCount; j++)
            {
                var v = Sheet.GetValue(j);

            }
        }


        file.Close();

        //Begin Config init, cause if start init on useing, maybe too long time
        var configScriptType = Type.GetType("Config." + _configScripName);
        if (configScriptType == null)
        {
            Debug.LogError($"Excel Reader Error :: doesn't find name is Config.{_configScripName}'s script");
        }
        else
        {
            PropertyInfo propertyInfo = configScriptType.GetProperty("Init");
            MethodInfo methodInfo = propertyInfo.GetGetMethod();
            object configInstance = Activator.CreateInstance(configScriptType);
            methodInfo.Invoke(configInstance, null);
        }
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
              "         }\n" +
              "\n" +
              $"         private static {_configScripName} _init = null;\n" +
              $"         public static {_configScripName} Init\n" +
              "         {\n" +
              "             get\n" +
              "             {\n" +
              "                 if (_init == null)\n" +
              $"                     _init = new {_configScripName}();\n" +
              "                 return _init;\n" +
              "             }\n" +
              "         }\n" +
              "\n" +
              "         //key is Data's ID\n" +
              $"        private Dictionary<int, {_configName}Data> _dataContainer = null;\n" +
              "\n" +
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
        initMethodArgsString.Append("         #pragma warning disable CS0414\n");
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
        initMethodArgsString.Append("         #pragma warning restore CS0414\n\n");
        _dataClass.Append( initMethodArgsString.ToString() +
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