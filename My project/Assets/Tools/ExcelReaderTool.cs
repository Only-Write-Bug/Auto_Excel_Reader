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
    [MenuItem("Tools/Excel Reader")]
    public static void ExcelReader()
    {
        var curDirectory = Directory.GetCurrentDirectory();

        var readFloder = Directory.GetDirectories(curDirectory, "Excel");
        var excelContainer = Directory.GetFiles(readFloder[0], "*.xlsx");

        var writeFloder = Directory.GetDirectories(curDirectory, "Config");

        Debug.LogError(excelContainer.Length);
    }
}
