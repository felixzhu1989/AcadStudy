using Autodesk.AutoCAD.Interop.Common;
using System;

namespace AcadStudy
{
    class Program
    {
        static void Main(string[] args)
        {
            //DrawingStudy drawing = new DrawingStudy();
            //drawing.CreateLine();
            //drawing.CreateSpline();

            //new Program().AcadMethodDemo();

            //BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
            //------测试代码---------------------
            CodeFromVBA demo = new CodeFromVBA();

            double[] startPoint = { 50, 50, 0 };
            double[] endPoint = { -50, 50, 0 };

            AcadLine line1 = demo.AddLineByPreDefine();
            AcadLine line2 = demo.AddLineByPoint(startPoint, endPoint);
            AcadLine line3 = demo.AddLineByXY(-50, 50, 0, 0);
            line1.Lineweight=ACAD_LWEIGHT.acLnWt030;
            line3.color = ACAD_COLOR.acRed;










            demo.Zoom();
            Console.WriteLine("绘制完成");
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
