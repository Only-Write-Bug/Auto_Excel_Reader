# Auto_Excel_Reader

Excel/Config/XML目录配置：在project/Excel2Config/ConfigurationInformation.txt中
备注：可以没有配置文件，如果需要自定义目录，在()中添加完整目录，不可使用中文字符，且不要添加多余字符（包括空格）
    如果没有配置文件，则会按照示例目录，创建文件夹
    存储Config的文件夹要放在Unity项目中的Assets目录下

在读表工具运行时，会有一些信息输出到日志文件中，日志文件在Excel2Config/ExcelReaderTool_WorkLogs中，按照工具启动时间命名
如果读表工具出现Bug，会在控制台停住，也就是Error，也会在控制台输出
日志文件不会自动清理，可以自行清理

务必按照规则填写，表格中除了备注（也不建议写中文），不要出现中文字符
支持CS中大部分常用基础类型及数组(数组以英文逗号分割)
支持自定义类型如下：
1.Vecotr3 :: [x|y|z]
2.Vector3[] :: [x|y|z],[x|y|z],……
3.Vecotr2 :: [x|y]
4.Vector2[] :: [x|y],[x|y],……

基本功能：
1.读取Excel，生成相应Config代码，在需要使用Config数据时，在脚本顶端Using Config，使用{ExcelName}Config.Init.*即可

期望功能：
1.Dirty Read :: 使用脏标记，减少对未修改Excel的读取和输出，以达到减少读表时间，增加工作效率的目的 (Open)

使用方法：
1.把项目中Excel2Config挪入Unity项目目录中（与Assets平级）
2.将Assets/Tools中的脚本放到自己项目的文件夹中
3.在Assets中创建Config文件夹，并将该路径写入上述配置文件中的指定位置
4.运行Excel2Config文件夹中的.exe快捷方式
5.Unity编辑器/菜单栏/Tools/ExcelReader/InitAllConfig，或在代码中直接调用ExcelReaderTool.InitAllConfig(),可以提前初始化Config；也可以直接调用Config.Init.Awake()进行单个congfig初始化；Config.Init.GetDataClassByID()，注：不建议未初始化就使用该方法

每个config都是静态的