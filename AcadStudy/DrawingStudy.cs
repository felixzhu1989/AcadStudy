using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace AcadStudy
{
    
    public class DrawingStudy
    {
        private AcadApplication acadApp = AcadConn.GetAcadApplication();
        private AcadDocument acadDoc;
        private AcadModelSpace moSpace;

        public DrawingStudy()
        {
            if (acadApp != null)
            {
                Console.WriteLine(acadApp.Name + " : " + acadApp.Version + "准备就绪");
            }
            else
            {
                Console.WriteLine("未连接到CAD");
                return;
            }
            acadDoc = acadApp.ActiveDocument;
            moSpace = acadDoc.ModelSpace;
        }


        public void CreateLine()
        {
            double[] startPoint = { 1, 1, 0 };
            double[] endPoint = { 5, 5, 0 };
            AcadLine line = moSpace.AddLine(startPoint, endPoint);
            Console.WriteLine(line.EntityName);
            acadApp.ZoomExtents();
            if (moSpace.Count != 0)
            {
                AcadEntity acadEntity = moSpace.Item(0);
                Console.WriteLine(acadEntity.GetType());
            }
        }

        public void CreateSpline()
        {
            double[] fitPoint = { 0, 0, 0, 5, 5, 0, 10, 0, 0 };
            double[] startTan = { 0.5, 0.5, 0 };
            double[] endTan = { 0.5, 0.5, 0 };

            AcadSpline spline = moSpace.AddSpline(fitPoint, startTan, endTan);
            Console.WriteLine(spline.EntityName);
            acadApp.ZoomExtents();
        }





    }
}
