# Auto_Excel_Reader

读表路径：Project Path/Excel
输出路径：Project Path/Assets/Config

支持类型基础类型及数组(数组以英文逗号分割)
支持自定义类型如下：
1.Vecotr3 :: [x,y,z]
2.Vector3[] :: [x,y,z],[x,y,z],……
3.Vecotr2 :: [x,y]
4.Vector2[] :: [x,y],[x,y],……

基本功能：
1.读取Excel，生成相应Config代码，在需要使用Config数据时，在脚本顶端Using Config，使用{ExcelName}Config.Init即可

期望功能：
1.Dirty Read :: 使用脏标记，减少对未修改Excel的读取和输出，以达到减少读表时间，增加工作效率的目的 (Open)