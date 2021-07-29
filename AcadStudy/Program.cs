using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Security.Policy;

namespace AcadStudy
{
    class Program
    {
        static void Main(string[] args)
        {
            //DrawingStudy drawing = new DrawingStudy();
            //drawing.CreateLine();
            //drawing.CreateSpline();

            AcadMethod method=new AcadMethod();
            method.DrawLineLeft(0, 0, 0, 600, true);
            method.DrawLineRight(200, 0, 200, 600, true);
            method.DrawLineUp(0, 600, 200, 600, true);
            method.DrawLineDown(0, 0, 200, 0, true);
            method.DrawRectangle(100, 100, 100, 200, true);
            method.AddBlock(0, 0, "C:\\Users\\Administrator\\Desktop\\Drawing1.dwg");


            Console.WriteLine("绘制完成");
            Console.ReadKey();
        }


    }
}
