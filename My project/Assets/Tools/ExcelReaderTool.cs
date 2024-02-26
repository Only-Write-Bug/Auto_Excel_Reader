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

public class ExcelReaderTool : MonoBehaviour
{
    private static string _readFloderPath = null;
    private static string _writeFloderPath = null;

    private static string _configName = null;
    private static StringBuilder _dataClass = new StringBuilder();

    [MenuItem("Tools/Excel Reader")]
    public static void ExcelReader()
    {
        //Get project path
        var curDirectory = Directory.GetCurrentDirectory();

        if (ExcelReaderTool._readFloderPath == null)
            _readFloderPath = Directory.GetDirectories(curDirectory, "Excel")[0];
        if (ExcelReaderTool._writeFloderPath == null)
            _writeFloderPath = Directory.GetDirectories(curDirectory + "/Assets", "Config")[0];

        ExcelReaderTool.Work();
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
                Debug.LogError($"{fileStream} convert to Excel Sheet fail£¡");
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

        foreach (var config in configFilesNameContainer)
        {
            File.Delete(config);
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

        //befor WriteDataClass(), Init Data
        string[][] dataInfo = new string[3][];
        for (int i = 0; i < 3; i++)
        {
            dataInfo[i] = new string[columnCount];
            Sheet.Read();
            for (int j = 0; j < columnCount; j++)
            {
                dataInfo[i][j] = Sheet.GetValue(j)?.ToString();
            }
        }

        WriteDataClass(dataInfo);

        InitConfigCS(file);

        for (int i = 3; i < rowCount; i++)
        {
            Sheet.Read();
            for (int j = 0; j < columnCount; j++)
            {
                var v = Sheet.GetValue(j);
            }
        }

        file.Close();
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
            "using System.Collections.Generic;\n\n";

        byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(data);
        file.Write(byteArray, 0, byteArray.Length);
    }

    private static void WriteStaticConfig(FileStream file)
    {
        StringBuilder nameSB = new StringBuilder();
        GetFileName(file.Name, nameSB);
        string name = nameSB.ToString();

        string data = "namespace Config\n" +
              "{\n" +
              $"     public class {name}\n" +
              "     {\n" +
              "         private " + name + "() {}\n" +
              "\n" +
              $"         private static {name} _init = null;\n" +
              $"         public static {name} Init\n" +
              "         {\n" +
              "             get\n" +
              "             {\n" +
              "                 if (_init == null)\n" +
              $"                     _init = new {name}();\n" +
              $"                _init.OnConfigInit();\n" +
              "                 return _init;\n" +
              "             }\n" +
              "         }\n" +
              "\n" +
              "         //key is Data's ID\n" +
              $"        private Dictionary<int, {_configName}Data> _dataContainer = new Dictionary<int, {_configName}Data>(8);\n" +
              "\n" +
              "         private void OnConfigInit()\n" +
              "         {\n" +
              "             \n" +
              "         }\n" +
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

            if (dataInfo[1][i] == null || dataInfo[2][i] == null)
            {
                Debug.LogError("Excel Reader Error :: Excel data incomplete information, please check your excel stats name or type is incomplete?");
                return;
            }

            _dataClass.Append(
                $"          public {dataInfo[2][i]} {dataInfo[1][i]};\n\n");
        }

        _dataClass.Append(
              "     }\n");

    }
}
