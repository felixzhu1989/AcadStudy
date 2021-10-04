using System;
using System.Windows.Forms;

namespace AcadStudy
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        /// [STAThread]
        static void Main()
        {


            Console.WriteLine("Connecting...");


            //AcadMethodTest();

            Chapter chapter = new Chapter();
            //chapter.P1to4();
            //chapter.P5();
            //chapter.P6();
            //chapter.P7();
            //chapter.P8_1();
            //chapter.P8_2();
            //chapter.P9to10();
            //chapter.P11to12();
            //chapter.P13();
            //chapter.P14toP15();
            //chapter.P16();
            //chapter.P17to18();
            //chapter.P19();
            //chapter.P20();
            //chapter.P21();
            //chapter.P22();
            //chapter.P23();
            //chapter.P24();
            //chapter.P25();
            //第26-28课，文字替换
            //RunWindow();//属性-应用程序-输出：Windows应用程序
            //chapter.P29();
            //chapter.P30();
            //chapter.P31();
            //chapter.P32to40();
            //chapter.P41();
            //chapter.P42();
            //chapter.P43();
            //chapter.P44_1();
            //chapter.P44_2();
            //chapter.P45();
            //chapter.P46();
            //chapter.P47();
            //chapter.P48();
            //chapter.P49();
            //chapter.P50();
            //chapter.P51();
            //chapter.P52();
            //chapter.P53();
            //chapter.P54();
            //chapter.P55();
            //chapter.P56();
            //chapter.P57();
            //chapter.P58();
            chapter.P59();


            Console.WriteLine("Drawing completed, please close the window！\n------------------------------------------");
            Console.ReadKey();//属性-应用程序-输出：控制台应用程序
        }


        /// <summary>
        /// 显示窗体
        /// </summary>
        public static void RunWindow()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmMain());
        }

        /// <summary>
        /// AcadMethod测试代码
        /// </summary>
        private static void AcadMethodTest()
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
