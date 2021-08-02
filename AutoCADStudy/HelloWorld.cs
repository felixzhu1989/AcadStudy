using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;


namespace AutoCADStudy
{
    //https://github.com/hongfengzhu/AutoCADStudy
    //打开CAD，输入NETLOAD，加载生成的dll类库，输入如下命令Hello
    //注意CAD2008，是.NET framework 3.5,CPU X86


    public class HelloWorld
    {
        [CommandMethod("Hello")]
        public void Hello()
        {
            //命令行对象
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Hello World!");
        }


    }
}
