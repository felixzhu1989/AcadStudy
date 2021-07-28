using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Interop;

namespace AcadStudy
{
    public class AcadSingleton
    {
        private static AcadApplication acadApp;
        /// <summary>
        /// 连接CAD程序
        /// </summary>
        /// <returns></returns>
        public static AcadApplication GetAcadApplication()
        {
            if (acadApp==null)
            {
                acadApp = Activator.CreateInstance(Type.GetTypeFromProgID("AutoCAD.Application")) as AcadApplication;
                acadApp.Visible = true;
                return acadApp;
            }
            return acadApp;
        }


    }
}
