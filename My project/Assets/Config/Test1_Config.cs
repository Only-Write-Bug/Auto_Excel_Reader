using UnityEngine;
using System;
using System.Xml;
using System.IO;
using System.Collections.Generic;

namespace Config
{
     public class Test1Config
     {
        const string XMLPath = "F:\\UnityProject\\Auto_Excel_Reader\\My project\\Excel2Config\\XML";

         public Test1Config() 
         {
             this._dataContainer = new Dictionary<int, Test1ConfigData>(8);
         }

         private static Test1Config _init = null;
         public static Test1Config Init
         {
             get
             {
                 if (_init == null)
                 {
                    _init = new Test1Config();
                    _init.OnCreate();
                 }
                 return _init;
             }
         }

         //key is Data's ID
        private Dictionary<int, Test1ConfigData> _dataContainer = null;

         private void OnCreate()
         {
             string xmlFileName = "Test1Config_XML.xml";
             string filePath = Path.Combine(XMLPath, xmlFileName);
             if(File.Exists(filePath))
             {
                 XmlDocument xml = new XmlDocument();
                 xml.Load(filePath);

                 XmlNode root = xml.DocumentElement;

                 foreach(XmlNode sonNode in root.ChildNodes)
                 {
                     Debug.Log(sonNode.Name);
                     foreach(XmlNode grandsonNode in sonNode.ChildNodes)
                     {
                         if(grandsonNode.Name == "#text")
                             continue;
                         Debug.Log(grandsonNode.Name);
                     }
                 }
             }
             else
             {
                 Debug.LogError("Excel Reader Error :: Test1Config_XML.xml is not find;");
             }
         }

     }

     public class Test1ConfigData
     {
          public Test1ConfigData(int? args_ID, string args_name, int? args_age, float? args_hight, bool? args_sex, Vector3? args_location, Vector3[] args_TestValue)
          {
               this._ID = args_ID ?? int.MinValue;
               this._name = args_name ?? null;
               this._age = args_age ?? int.MinValue;
               this._hight = args_hight ?? float.MinValue;
               this._sex = args_sex ?? false;
               this._location = args_location ?? Vector3.zero;
               this._TestValue = args_TestValue ?? null;
          }

          private int _ID;
          public int Get_ID
          {
               get { return _ID; }
          }

          private string _name;
          public string Get_name
          {
               get { return _name; }
          }

          private int _age;
          public int Get_age
          {
               get { return _age; }
          }

          private float _hight;
          public float Get_hight
          {
               get { return _hight; }
          }

          private bool _sex;
          public bool Get_sex
          {
               get { return _sex; }
          }

          private Vector3 _location;
          public Vector3 Get_location
          {
               get { return _location; }
          }

          private Vector3[] _TestValue;
          public Vector3[] Get_TestValue
          {
               get { return _TestValue; }
          }


     }

}
