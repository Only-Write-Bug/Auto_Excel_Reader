using UnityEngine;
using System;
using System.Collections.Generic;

namespace Config
{
     public class Test1Config
     {
         private Test1Config() {}

         private static Test1Config _init = null;
         public static Test1Config Init
         {
             get
             {
                 if (_init == null)
                     _init = new Test1Config();
                _init.OnConfigInit();
                 return _init;
             }
         }

         //key is Data's ID
        private Dictionary<int, Test1Data> _dataContainer = new Dictionary<int, Test1Data>(8);

         private void OnConfigInit()
         {
             
         }
     }

     public class Test1Data
     {
          //唯一标识
          internal int _ID;
          public int ID
          {
              get { return this._ID; }
          }

          //名字
          internal string _name;
          public string name
          {
              get { return this._name; }
          }

          //how old
          internal int _age;
          public int age
          {
              get { return this._age; }
          }

          internal float _hight;
          public float hight
          {
              get { return this._hight; }
          }

          internal bool _sex;
          public bool sex
          {
              get { return this._sex; }
          }

          internal Vector3 _location;
          public Vector3 location
          {
              get { return this._location; }
          }

          internal Vector3[] _TestValue;
          public Vector3[] TestValue
          {
              get { return this._TestValue; }
          }

     }

}
