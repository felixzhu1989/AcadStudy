using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace AcadStudy
{
    
    public class AcadTest
    {
        private AcadApplication acadApp = AcadConn.GetAcadApplication();
        private AcadDocument acadDoc;
        private AcadModelSpace modelSpace;

        public AcadTest()
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
            modelSpace = acadDoc.ModelSpace;
        }


        public void CreateLine()
        {
            double[] startPoint = { 1, 1, 0 };
            double[] endPoint = { 5, 5, 0 };
            AcadLine line = modelSpace.AddLine(startPoint, endPoint);
            Console.WriteLine(line.EntityName);
            if (modelSpace.Count != 0)
            {
                AcadEntity acadEntity = modelSpace.Item(0);
                Console.WriteLine(acadEntity.GetType());
            }
            acadApp.ZoomExtents();
        }

        public void CreateSpline()
        {
            double[] fitPoint = { 0, 0, 0, 5, 5, 0, 10, 0, 0 };
            double[] startTan = { 0.5, 0.5, 0 };
            double[] endTan = { 0.5, 0.5, 0 };

            AcadSpline spline = modelSpace.AddSpline(fitPoint, startTan, endTan);
            Console.WriteLine(spline.EntityName);
            acadApp.ZoomExtents();
        }





    }
}
