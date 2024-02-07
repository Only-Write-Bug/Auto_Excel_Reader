using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using System.Data;
using System;
using ExcelDataReader;

public class ExcelReaderTool : MonoBehaviour
{
    private static String _readFloder = null;
    private static String _writeFloder = null;

    [MenuItem("Tools/Excel Reader")]
    public static void ExcelReader()
    {
        //Get project path
        var curDirectory = Directory.GetCurrentDirectory();

        if (ExcelReaderTool._readFloder == null)
            _readFloder = Directory.GetDirectories(curDirectory, "Excel")[0];
        if (ExcelReaderTool._writeFloder == null)
            _writeFloder = Directory.GetDirectories(curDirectory, "Config")[0];

        ExcelReaderTool.Work();
    }

    private static void Work()
    {
        //Get all excel files
        var excelFilesNameContainer = Directory.GetFiles(_readFloder, "*.xlsx");

        foreach (var fileName in excelFilesNameContainer)
        {
            string excelFilePath = fileName;
            FileStream excelFile = File.OpenRead(excelFilePath);
            IExcelDataReader firstSheet = ExcelReaderFactory.CreateReader(excelFile);

            if (firstSheet == null)
            {
                Debug.LogError($"{fileName} convert to Excel Sheet fail£¡");
                continue;
            }

            ExcelReaderTool.ReadExcelSheet(firstSheet);
        }
    }

    private static void ReadExcelSheet(IExcelDataReader Sheet)
    {
        //Number of valid columns on worksheet
        int columnCount = Sheet.FieldCount;
        //Number of valid rows on worksheet
        int rowCount = Sheet.RowCount;

        
    }
}
