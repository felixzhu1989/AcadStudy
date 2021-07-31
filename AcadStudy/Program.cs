using Autodesk.AutoCAD.Interop.Common;
using System;

namespace AcadStudy
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Connecting...");

            //new Program().AcadMethodDemo();

            Chapter chapter = new Chapter();
            //chapter.P1_4();
            //chapter.P5();
            //chapter.P6();
            chapter.P7();












            Console.WriteLine("Drawing work done！\n------------------");
            Console.ReadKey();
        }

        /// <summary>
        /// AcadMethod测试代码
        /// </summary>
        private void AcadMethodDemo()
        {
            AcadMethod method = new AcadMethod();
            method.DrawLineLeft(0, 0, 0, 600, true);
            method.DrawLineRight(200, 0, 200, 600, true);
            method.DrawLineUp(0, 600, 200, 600, true);
            method.DrawLineDown(0, 0, 200, 0, true);
            method.DrawRectangle(100, 100, 100, 200, true);
            method.AddBlock(0, 0, "C:\\Users\\Administrator\\Desktop\\Drawing1.dwg");
        }





    }
}
