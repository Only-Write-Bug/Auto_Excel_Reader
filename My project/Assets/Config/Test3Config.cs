using UnityEngine;
using System;
using System.Collections.Generic;

namespace Config
{
     public class Test3Config
     {
         private Test3Config() {}

         private static Test3Config _init = null;
         public static Test3Config Init
         {
             get
             {
                 if (_init == null)
                     _init = new Test3Config();
                _init.OnConfigInit();
                 return _init;
             }
         }

         //key is Data's ID
        private Dictionary<int, Test3Data> _dataContainer = new Dictionary<int, Test3Data>(8);

         private void OnConfigInit()
         {
             
         }
     }

     public class Test3Data
     {
     }

}
