using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Autodesk.AutoCAD.Interop;

namespace AcadStudy
{
    /// <summary>
    /// 参考：https://blog.csdn.net/shengshaohua/article/details/8363769
    /// </summary>
    public class AcadConn : IDisposable
    {
        private static AcadApplication acadApp;
        private static bool initialized;
        private bool disposed;

        public static AcadApplication GetAcadApplication()
        {
            try
            {
                // Upon creation, attempt to retrieve running instance
                acadApp = (AcadApplication)Marshal.GetActiveObject("AutoCAD.Application");
                //acadApp.Documents.Add("");//添加一个新的drawing
            }
            catch
            {
                try
                {
                    // Create an instance and set flag to indicate this
                    acadApp = new AcadApplication();
                    initialized = true;
                    acadApp.Visible = true;
                    return acadApp;
                }
                catch
                {
                    throw;
                }
            }
            acadApp.Visible = true;
            return acadApp;
        }




        // If the user doesn't call Dispose, the
        // garbage collector will upon destruction
        ~AcadConn()
        {
            try { Dispose(false); }
            catch { }
        }

        public AcadApplication Application
        {
            get
            {
                // Return our internal instance of AutoCAD
                return acadApp;
            }
        }

        // This is the user-callable version of Dispose.
        // It calls our internal version and removes the
        // object from the garbage collector's queue.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        // This version of Dispose gets called by our
        // destructor.
        protected virtual void Dispose(bool disposing)
        {

            // If we created our AutoCAD instance, call its
            // Quit method to avoid leaking memory.

            if (!this.disposed && initialized)
                acadApp.Quit();
            disposed = true;
        }
    }
}
