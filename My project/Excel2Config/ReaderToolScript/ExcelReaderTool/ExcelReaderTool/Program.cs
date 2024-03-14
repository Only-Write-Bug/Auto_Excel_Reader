using System;

namespace ExcelReaderTool
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelReaderTool.Init.CheckWorkPath();
            if(ExcelReaderTool.Init.Get_ErrorCount > 0)
            {
                Console.WriteLine("Excel Reader Tool :: have Error, please check log file!!!");
                Console.ReadKey();
            }
#if DEBUG
            Console.ReadKey();
#endif
        }
    }
}
