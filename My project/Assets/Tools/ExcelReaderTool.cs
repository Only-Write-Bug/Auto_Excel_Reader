using UnityEngine;
using UnityEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Config
{
    public static class ExcelReadTool
    {
        [MenuItem("Tools/ExcelReader/InitAllConfig")]
        public static void InitAllConfig()
        {
            Debug.Log("Excel Reader start work");
            var startTime = DateTime.Now;

            Type[] types = Assembly.GetExecutingAssembly().GetTypes();

            var classesWithAttribute = types
                .Where(type => type.GetCustomAttributes(typeof(ConfigAttribute), true).Any());

            foreach (var classType in classesWithAttribute)
            {
                PropertyInfo initProperty = classType.GetProperty("Init", BindingFlags.Public | BindingFlags.Static);
                object instance = initProperty?.GetValue(null);

                MethodInfo method = classType.GetMethod("Awake");
                method?.Invoke(instance, null);
            }

            Debug.Log("Excel Reader Work is Over, Work time :: " + (DateTime.Now - startTime));
        }

        internal static T XmlDataConvert<T>(string type, string value, T defaultValue)
        {
            if (value.ToLower() == "null")
                return defaultValue;

            T tmpValue = defaultValue;

            switch (type)
            {
                case "int":
                    if (int.TryParse(value, out var intValue))
                        tmpValue = (T)(object)intValue;
                    break;
                case "short":
                    if (short.TryParse(value, out var shortValue))
                        tmpValue = (T)(object)shortValue;
                    break;
                case "long":
                    if (long.TryParse(value, out var longValue))
                        tmpValue = (T)(object)longValue;
                    break;
                case "float":
                    if (float.TryParse(value, out var floatValue))
                        tmpValue = (T)(object)floatValue;
                    break;
                case "double":
                    if (double.TryParse(value, out var doubleValue))
                        tmpValue = (T)(object)doubleValue;
                    break;
                case "bool":
                    if (bool.TryParse(value, out var boolValue))
                        tmpValue = (T)(object)boolValue;
                    break;
                case "char":
                    if (char.TryParse(value, out var charValue))
                        tmpValue = (T)(object)charValue;
                    break;
                case "Vector2":
                case "UnityEngine.Vector2":
                    value = value.Substring(1, value.Length - 2);
                    string[] vector2Components = value.Split('|');
                    if (vector2Components.Length == 2 &&
                        float.TryParse(vector2Components[0], out float x) &&
                        float.TryParse(vector2Components[1], out float y))
                    {
                        tmpValue = (T)(object)new Vector2(x, y);
                    }

                    break;
                case "Vector3":
                case "UnityEngine.Vector3":
                    value = value.Substring(1, value.Length - 2);
                    string[] vector3Components = value.Split('|');
                    if (vector3Components.Length == 3 &&
                        float.TryParse(vector3Components[0], out float x1) &&
                        float.TryParse(vector3Components[1], out float y1) &&
                        float.TryParse(vector3Components[2], out float z))
                    {
                        tmpValue = (T)(object)new Vector3(x1, y1, z);
                    }

                    break;
            }

            return tmpValue;
        }

        internal static T[] XMLDataConvertArray<T>(string type, string value, T defaultValue)
        {
            var modelType = type.Substring(0, type.Length - 2);

            var valueType = GetArrayModelType(type);

            if (valueType != null)
            {
                value = value.Substring(1, value.Length - 2);
                string[] arrayValues = value.Split(',');
                var list = new List<T>();

                foreach (string arrayValue in arrayValues)
                {
                    T elementValue = XmlDataConvert(modelType, arrayValue.Trim(), defaultValue);
                    list.Add(elementValue);
                }

                // Array array = list.ToArray();
                // Array convertedArray = Array.CreateInstance(valueType, array.Length);
                // Array.Copy(array, convertedArray, array.Length);
                return list.ToArray();
            }

            return null;
        }

        internal static Type GetArrayModelType(string type)
        {
            var modelType = type.Substring(0, type.Length - 2);

            Assembly otherAssembly = null;
            bool isNotCSNativeType = false;

            switch (modelType)
            {
                case "Vector3":
                case "Vector2":
                    otherAssembly = typeof(UnityEngine.Vector2).Assembly;
                    isNotCSNativeType = true;
                    modelType = "UnityEngine." + modelType;
                    break;
            }

            modelType = !isNotCSNativeType ? NativeTypeCorrection(modelType) : modelType;

            return isNotCSNativeType ? otherAssembly.GetType(modelType) : Type.GetType(modelType);
        }

        private static string NativeTypeCorrection(string type)
        {
            switch (type)
            {
                case "int":
                    return "System.Int32";
                case "bool":
                    return "System.Boolean";
                case "byte":
                    return "System.Byte";
                case "sbyte":
                    return "System.SByte";
                case "char":
                    return "System.Char";
                case "decimal":
                    return "System.Decimal";
                case "double":
                    return "System.Double";
                case "float":
                    return "System.Single";
                case "long":
                    return "System.Int64";
                case "ulong":
                    return "System.UInt64";
                case "short":
                    return "System.Int16";
                case "ushort":
                    return "System.UInt16";
                case "string":
                    return "System.String";
                case "object":
                    return "System.Object";
                default:
                    return null;
            }
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