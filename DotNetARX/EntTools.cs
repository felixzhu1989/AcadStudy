using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace DotNetARX
{
	//https://archive.codeplex.com
	
    /// <summary>
    /// 工具类
    /// </summary>
    public static class EntTools
    {
        /// <summary>
        /// 将实体添加到模型空间(扩展方法)
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到模型空间中的实体</returns>
        public static ObjectId AddToModelSpace(this Database db, Entity ent)
        {
            ObjectId entId;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(
                    db.BlockTableId, OpenMode.ForRead);//块表
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(
                    bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);//块表记录
                entId=btr.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent, true);//将对象添加到事物中
                trans.Commit();//提交事务
            }
            return entId;
        }
        /// <summary>
        /// 移动实体
        /// </summary>
        /// <param name="id">实体对象</param>
        /// <param name="sourcePt">起点</param>
        /// <param name="targetPt">终点</param>
        public static void Move(this ObjectId id, Point3d sourcePt, Point3d targetPt)
        {
            //构建用于移动实体的矩阵
            Vector3d vector = targetPt.GetVectorTo(sourcePt);
            Matrix3d mt=Matrix3d.Displacement(vector);
            //以写的方式打开id表示的实体对象
            Entity ent = (Entity) id.GetObject(OpenMode.ForWrite);
            ent.TransformBy(mt);//对实体实施移动
            ent.DowngradeOpen();//为防止错误将实体切换为读的状态
        }
        /// <summary>
        /// 重载移动函数
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="sourcePt">起点</param>
        /// <param name="targetPt">终点</param>
        public static void Move(this Entity ent, Point3d sourcePt, Point3d targetPt)
        {
            if (ent.IsNewObject)//如果还是未添加到数据库中的新实体
            {
                Vector3d vector = targetPt.GetVectorTo(sourcePt);
                Matrix3d mt = Matrix3d.Displacement(vector);
                ent.TransformBy(mt);
            }
            else
            {
                //如果已经是添加到数据库中的实体，使用上面的扩展方法移动
                ent.ObjectId.Move(sourcePt,targetPt);
            }
        }




    }
}
