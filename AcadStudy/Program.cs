using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Security.Policy;

namespace AcadStudy
{
    class Program
    {
        private static AcadApplication acadApp;
        static void Main(string[] args)
        {
            acadApp = AcadSingleton.GetAcadApplication();
            if (acadApp != null)
            {
                Console.WriteLine(acadApp.Name + " : " + acadApp.Version + "准备就绪");
            }
            else
            {
                Console.WriteLine("未连接到CAD");
                return;
            }

            AcadDocument acadDoc = acadApp.ActiveDocument;
            AcadModelSpace moSpace = acadDoc.ModelSpace;
            double[] startPoint = { 1, 1, 0 };
            double[] endPoint = { 5, 5, 0 };
            AcadLine line = moSpace.AddLine(startPoint, endPoint);
            Console.WriteLine(line.EntityName);

            double[] fitPoint = {0, 0, 0, 5, 5, 0, 10, 0, 0};
            double[] startTan ={ 0.5, 0.5, 0 };
            double[] endTan = { 0.5, 0.5, 0 };

            AcadSpline spline = moSpace.AddSpline(fitPoint, startTan, endTan);
            Console.WriteLine(spline.EntityName);

            acadApp.ZoomExtents();
            AcadEntity acadEntity = null;
            if (moSpace.Count!=0)
            {
                acadEntity = moSpace.Item(0);
                Console.WriteLine(acadEntity.GetType());
            }








            Console.WriteLine("绘制完成");
            Console.ReadKey();





        }
    }
}
