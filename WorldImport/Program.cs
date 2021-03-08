using System;

namespace WorldImport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ImportWorld world = new ImportWorld();
                world.Import(@"C:\Users\zengwang\Desktop\采购合同范本-修订稿-20210105\\3-采购合同---测试.docx");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            Console.WriteLine("Hello World!");
            Console.ReadLine();
        }
    }
}
