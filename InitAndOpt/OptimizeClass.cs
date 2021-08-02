using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace InitAndOpt
{
    public class OptimizeClass
    {
        [CommandMethod("OptCommand")]
        public void OptCommand()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            string fileName = @"D:\AutoCadDemo\AutoCADStudy\AutoCADStudy\bin\Debug\AutoCADStudy.dll";
            try
            {
                ExtensionLoader.Load(fileName); //载入上面地址中的程序集AutoCADStudy.dll
                ed.WriteMessage("\n" + fileName + "被载入，轻输入Hello测试！");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.Message);
            }
            finally
            {
                ed.WriteMessage("\n" +"程序执行完毕！");
            }
        }

        [CommandMethod("ChangeColor")]
        public void ChangeColor()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                //提示用户选择对象
                ObjectId id = ed.GetEntity("\n请选择要改变颜色的对象").ObjectId;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    //以写的方式打开对象
                    Entity ent = (Entity) trans.GetObject(id, OpenMode.ForWrite);
                    ent.ColorIndex = 100;//10red,100green
                    trans.Commit();//提交事务
                }

            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                switch (ex.ErrorStatus)
                {
                    case ErrorStatus.InvalidIndex:
                        ed.WriteMessage("\n输入的颜色值有误！");
                        break;
                    case ErrorStatus.InvalidObjectId:
                        ed.WriteMessage("\n未选择对象！");
                        break;
                        default:
                        ed.WriteMessage(ex.ErrorStatus.ToString());
                        break;
                }
            }
        }
    }
}
