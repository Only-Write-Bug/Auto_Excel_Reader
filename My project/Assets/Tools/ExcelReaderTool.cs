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
        var parentDirectory = Directory.GetParent(curDirectory).FullName;
        var lastParentDirectory = Directory.GetParent(parentDirectory).FullName;

        Debug.LogError(lastParentDirectory);
    }
}
