using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System;

namespace AcadStudy
{
    /*《ActiveX 和 VBA 参考》由明经通道翻译并提供：https://blog.csdn.net/weixin_46656590/article/details/105247189 
    下载地址：https://yunpan.360.cn/surl_yRE3CKKEqLk （提取码：c16e）
    BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
    */


    public class AcadHelper
    {
        private AcadApplication acadApp = AcadConn.GetAcadApplication();
        private AcadDocument acadDoc;
        private AcadModelSpace modelSpace;

        public AcadHelper()
        {
            Console.WriteLine("AutoCAD " + acadApp.Version + " Ready-->");

            acadDoc = acadApp.ActiveDocument;
            modelSpace = acadDoc.ModelSpace;
        }

        /// <summary>
        /// 将当前视图缩放到图形界限。
        /// </summary>
        public void Zoom()
        {
            acadApp.ZoomExtents();
        }
        /// <summary>
        /// 查看图元
        /// </summary>
        public void ShowEntity()
        {
            Object objEntity;
            Object pointPicked;
            acadDoc.Utility.GetEntity(out objEntity, out pointPicked, "请选择图元：");
            AcadEntity objAcadEntity = (AcadEntity)objEntity;
            Console.WriteLine(objAcadEntity.EntityName);
        }

        /// <summary>
        /// 绘制直线Demo
        /// </summary>
        /// <returns></returns>
        public AcadLine AddLineDemo()
        {
            double[] startPoint = { 0, 0, 0 };
            double[] endPoint = { 50, 50, 0 };
            AcadLine line = modelSpace.AddLine(startPoint, endPoint);
            return line;
        }
        /// <summary>
        /// 通过起始和结束点绘制直线
        /// </summary>
        /// <param name="startPoint">起始点</param>
        /// <param name="endPoint">结束点</param>
        /// <returns></returns>
        public AcadLine AddLineByPoint(double[] startPoint, double[] endPoint)
        {
            AcadLine line = modelSpace.AddLine(startPoint, endPoint);
            return line;
        }
        /// <summary>
        /// 通过起始和结束点坐标值绘制直线
        /// </summary>
        /// <param name="startX">起始点X坐标</param>
        /// <param name="stratY">起始点Y坐标</param>
        /// <param name="endX">结束点X坐标</param>
        /// <param name="endY">结束点Y坐标</param>
        /// <returns></returns>
        public AcadLine AddLineByXY(double startX, double stratY, double endX, double endY)
        {
            double[] startPoint = { startX, stratY, 0 };
            double[] endPoint = { endX, endY, 0 };
            AcadLine line = modelSpace.AddLine(startPoint, endPoint);
            return line;
        }
        /// <summary>
        /// 通过起点和相对点的直角坐标创建直线
        /// </summary>
        /// <param name="startPoint">开始点</param>
        /// <param name="endX">结束点X坐标</param>
        /// <param name="endY">结束点Y坐标</param>
        /// <returns></returns>
        public AcadLine AddLineByReXY(double[] startPoint, double endX, double endY)
        {
            double[] endPoint = { endX, endY, 0 };
            AcadLine line = modelSpace.AddLine(startPoint, endPoint);
            return line;
        }
        /// <summary>
        /// 通过起点和相对点的极坐标（角度和距离）创建直线
        /// </summary>
        /// <param name="startPoint">开始点</param>
        /// <param name="angle">角度</param>
        /// <param name="distance">距离</param>
        /// <returns></returns>
        public AcadLine AddLineByReAngle(double[] startPoint, double angle, double distance)
        {
            double[] endPoint = { startPoint[0] + distance * Math.Cos(angle), startPoint[1] + distance * Math.Sin(angle), 0 };
            AcadLine line = modelSpace.AddLine(startPoint, endPoint);
            return line;
        }
        /// <summary>
        /// 绘制轻量多段线Demo
        /// </summary>
        /// <returns></returns>
        public AcadLWPolyline AddLWPolylineDemo()
        {
            //两两构成一个点,这里有三点
            double[] point = { 0, 0, 100, 0, 100, 100 };
            if ((point.Length + 1) % 3 == 0)
            {
                Console.WriteLine("数组个数不匹配");
                return null;
            }
            else
            {
                AcadLWPolyline lwPolyline = modelSpace.AddLightWeightPolyline(point);
                lwPolyline.Closed = true;//闭合
                lwPolyline.ConstantWidth = 0.5;//线宽
                //AcadAcCmColor acColor = new AcadAcCmColorClass();//这种方式会报错
                AcadAcCmColor acColor = (AcadAcCmColor)acadDoc.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");
                acColor.ColorIndex = (AcColor)6;//注意类型强制转换
                lwPolyline.TrueColor = acColor;

                return lwPolyline;
            }
        }
        /// <summary>
        /// 绘制多段线Demo
        /// </summary>
        /// <returns></returns>
        public AcadPolyline AddPolylineDemo()
        {
            //三个数构成一个点,这里有三点
            double[] point = { 0, 0, 0, -100, 0, 0, -100, -100, 0 };
            if ((point.Length + 1) % 3 == 0)
            {
                Console.WriteLine("数组个数不匹配");
                return null;
            }
            else
            {
                AcadPolyline polyline = modelSpace.AddPolyline(point);
                polyline.Closed = true;
                polyline.color = ACAD_COLOR.acBlue;
                return polyline;
            }
        }
        /// <summary>
        /// 绘制圆Demo
        /// </summary>
        /// <returns></returns>
        public AcadCircle AddCircleDemo()
        {
            double[] centerPoint = { 0, 0, 0 };
            double radius = 50;
            AcadCircle circle = modelSpace.AddCircle(centerPoint, radius);
            return circle;
        }
        /// <summary>
        /// 通过圆心和半径绘制圆
        /// </summary>
        /// <param name="centerPoint">中心点</param>
        /// <param name="radius">半径</param>
        /// <returns></returns>
        public AcadCircle AddCircleByRadius(double[] centerPoint, double radius)
        {
            AcadCircle circle = modelSpace.AddCircle(centerPoint, radius);
            return circle;
        }
        /// <summary>
        /// 通过圆心和直径绘制圆
        /// </summary>
        /// <param name="centerPoint">中心点</param>
        /// <param name="diameter">直径</param>
        /// <returns></returns>
        public AcadCircle AddCircleByDiameter(double[] centerPoint, double diameter)
        {
            AcadCircle circle = modelSpace.AddCircle(centerPoint, diameter / 2);
            return circle;
        }
        /// <summary>
        /// 通过两点创建圆，两点为圆直径上两点
        /// </summary>
        /// <param name="firstPoint">第一点</param>
        /// <param name="secondPoint">第二点</param>
        /// <returns></returns>
        public AcadCircle AddCircleBy2Point(double[] firstPoint, double[] secondPoint)
        {
            double[] centerPoint = { (firstPoint[0] + secondPoint[0]) / 2, (firstPoint[1] + secondPoint[1]) / 2, 0 };
            double radius = Math.Sqrt(Math.Pow(firstPoint[0] - secondPoint[0], 2) + Math.Pow(firstPoint[1] - secondPoint[1], 2))/2;
            AcadCircle circle = modelSpace.AddCircle(centerPoint, radius);
            return circle;
        }

    }
}
