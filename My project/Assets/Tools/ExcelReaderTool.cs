using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using System.Data;
using System;
using ExcelDataReader;
using System.Text;

public class ExcelReaderTool : MonoBehaviour
{
    private static String _readFloderPath = null;
    private static String _writeFloderPath = null;

    [MenuItem("Tools/Excel Reader")]
    public static void ExcelReader()
    {
        //Get project path
        var curDirectory = Directory.GetCurrentDirectory();

        if (ExcelReaderTool._readFloderPath == null)
            _readFloderPath = Directory.GetDirectories(curDirectory, "Excel")[0];
        if (ExcelReaderTool._writeFloderPath == null)
            _writeFloderPath = Directory.GetDirectories(curDirectory, "Config")[0];

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
            Debug.Log($"Create new file :: {csFile.Name}");

            ExcelReaderTool.ReadExcelSheet_WriteCSFile(firstSheet, csFile);
        }
    }

    private static void ClearConfig()
    {
        //Get all config files
        var configFilesNameContainer = Directory.GetFiles(_writeFloderPath, "*.*");

        foreach(var config in configFilesNameContainer)
        {
            File.Delete(config);
        }
    }

    private static string FileNameFactory(string stream)
    {
        StringBuilder name = new StringBuilder();

        char[] charArray = stream.ToCharArray();
        foreach(char c in charArray)
        {
            if(c == '\\')
            {
                name.Clear();
                continue;
            }
            
            if(c == '.')
            {
                break;
            }

            name.Append(c);
        }
        name.Append("Config.cs");

        return _writeFloderPath + '\\' + name.ToString();
    }

    private static void ReadExcelSheet_WriteCSFile(IExcelDataReader Sheet, FileStream writeFile)
    {
        //Number of valid columns on worksheet
        int columnCount = Sheet.FieldCount;
        //Number of valid rows on worksheet
        int rowCount = Sheet.RowCount;

        for (int i = 0; i < rowCount; i++)
        {
            Sheet.Read();
            for(int j = 0; j < columnCount; j++)
            {
                var v = Sheet.GetValue(j);
                Debug.Log(v);
            }
        }

    }
}
