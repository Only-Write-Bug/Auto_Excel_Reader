using UnityEngine;
using System;
using System.Collections.Generic;

namespace Config
{
     public class Test2Config
     {
         private Test2Config() {}

         private static Test2Config _init = null;
         public static Test2Config Init
         {
             get
             {
                 if (_init == null)
                     _init = new Test2Config();
                _init.OnConfigInit();
                 return _init;
             }
         }

         //key is Data's ID
        private Dictionary<int, Test2Data> _dataContainer = new Dictionary<int, Test2Data>(8);

         private void OnConfigInit()
         {
             
         }
     }

     public class Test2Data
     {
     }

}
