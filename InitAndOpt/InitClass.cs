using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
//定义InitClass为程序入口点，而让AutoCad只会执行OptimizeClass
[assembly: ExtensionApplication(typeof(InitAndOpt.InitClass))]
[assembly: CommandClass(typeof(InitAndOpt.OptimizeClass))]


namespace InitAndOpt
{
    public class InitClass : IExtensionApplication
    {
        public void Initialize()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("程序开始初始化！");
        }

        public void Terminate()
        {
            System.Diagnostics.Debug.WriteLine("程序结束，你可以做一些清理工作，如关闭AutoCAD");
        }

        [CommandMethod("InitCommand")]
        public void InitCommand()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Test！");
        }

        

    }
}
