using UnityEngine;
using UnityEditor;
using System;
using System.Linq;
using System.Reflection;

namespace Config
{
    public static class TypeConversionTool
    {
        [MenuItem("Tools/ExcelReader/InitAllConfig")]
        public static void InitAllConfig()
        {
            var startTime = DateTime.Now;

            Type[] types = Assembly.GetExecutingAssembly().GetTypes();

            var classesWithAttribute = types
                .Where(type => type.GetCustomAttributes(typeof(ConfigAttribute), true).Any());

            foreach (var classType in classesWithAttribute)
            {
                PropertyInfo initProperty = classType.GetProperty("Init", BindingFlags.Public | BindingFlags.Static);
                object instance = initProperty.GetValue(null);

                MethodInfo method = classType.GetMethod("Awake");
                method.Invoke(instance, null);
            }
            
            Debug.Log("Excel Reader Work is Over, Work time :: " + (DateTime.Now - startTime));
        }
    }

    [AttributeUsage(AttributeTargets.Class)]
    public class ConfigAttribute : Attribute
    {
        public ConfigAttribute(string _className)
        {
            this.className = _className;
        }
        
        public string className = null;
    }
}