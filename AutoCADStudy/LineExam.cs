using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows.ToolPalette;
using DotNetARX;

namespace AutoCADStudy
{
    //https://www.bilibili.com/video/BV1x4411y79M

    
    public class LineExam
    {
        [CommandMethod("FirstLine")]
        public void FirstLine()
        {
            //图形数据库
            Database db = HostApplicationServices.WorkingDatabase;
            Point3d startPoint = new Point3d(0, 100, 1);
            Point3d endPoint = new Point3d(100, 100, 1);
            Line line = new Line(startPoint,endPoint);//直线对象

            EntTools.AddToModelSpace(db, line); //静态方法调用

            ////定义一个指向当前数据库的事务处理
            //using (Transaction trans = db.TransactionManager.StartTransaction())
            //{
            //    BlockTable bt = (BlockTable) trans.GetObject(
            //        db.BlockTableId, OpenMode.ForRead);//块表
            //    BlockTableRecord btr = (BlockTableRecord) trans.GetObject(
            //        bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);//块表记录
            //    btr.AppendEntity(line);
            //    trans.AddNewlyCreatedDBObject(line,true);//将对象添加到事物中
            //    trans.Commit();//提交事务
            //}
        }

        [CommandMethod("SecondLine")]
        public void SecondLine()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Point3d startPoint = new Point3d(0, 100, 1);
            Point3d endPoint = new Point3d(0, 200, 1);
            Line line = new Line(startPoint, endPoint);
            db.AddToModelSpace(line);//调用扩展方法
            
        }

    }
}
