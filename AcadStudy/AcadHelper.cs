using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace AcadStudy
{
    /*《ActiveX 和 VBA 参考》由明经通道翻译并提供：https://blog.csdn.net/weixin_46656590/article/details/105247189 
    下载地址：https://yunpan.360.cn/surl_yRE3CKKEqLk （提取码：c16e）
    BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
    代码由VBA的逻辑转化而来


    查看图元，CAD中lisp命令：(setq en_data(entget(car(entsel))))
    */


    public class AcadHelper
    {
        public AcadApplication acadApp = AcadConn.GetAcadApplication();
        public AcadDocument acadDoc;
        public AcadModelSpace modelSpace;

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
            Object pickedPoint;
            AcadEntity entity;
            do
            {
                try
                {
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
                    entity = (AcadEntity)objEntity;
                    Console.WriteLine("EntityName:" + entity.EntityName);
                    Console.WriteLine("EntityType:" + entity.EntityType);
                    Console.WriteLine("Handle:" + entity.Handle);
                    Console.WriteLine("ObjectID:" + entity.ObjectID);
                    Console.WriteLine("ObjectName:" + entity.ObjectName);
                }
                catch (Exception)
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") return;
                }
            } while (true);
        }

        #region 绘制图元

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
            double[] endPoint =
                {startPoint[0] + distance * Math.Cos(angle), startPoint[1] + distance * Math.Sin(angle), 0};
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
            double[] points = { 0, 0, 100, 0, 100, 100 };
            if ((points.Length + 1) % 3 == 0)
            {
                Console.WriteLine("数组个数不匹配");
                return null;
            }
            else
            {
                AcadLWPolyline lwPolyline = modelSpace.AddLightWeightPolyline(points);
                lwPolyline.Closed = true; //闭合
                lwPolyline.ConstantWidth = 0.5; //线宽
                //AcadAcCmColor acColor = new AcadAcCmColorClass();//这种方式会报错
                AcadAcCmColor acColor = (AcadAcCmColor)acadDoc.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");
                acColor.ColorIndex = (AcColor)6; //注意类型强制转换
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
            double[] points = { 0, 0, 0, -100, 0, 0, -100, -100, 0 };
            if ((points.Length + 1) % 3 == 0)
            {
                Console.WriteLine("数组个数不匹配");
                return null;
            }
            else
            {
                AcadPolyline polyline = modelSpace.AddPolyline(points);
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
        /// <param name="point1">第一点</param>
        /// <param name="point2">第二点</param>
        /// <returns></returns>
        public AcadCircle AddCircleBy2Point(double[] point1, double[] point2)
        {
            double[] centerPoint = { (point1[0] + point2[0]) / 2, (point1[1] + point2[1]) / 2, 0 };
            double radius = Math.Sqrt(Math.Pow(point1[0] - point2[0], 2) + Math.Pow(point1[1] - point2[1], 2)) / 2;
            AcadCircle circle = modelSpace.AddCircle(centerPoint, radius);
            return circle;
        }

        /// <summary>
        /// 通过三点创建圆，几何法
        /// </summary>
        /// <param name="point1">点1</param>
        /// <param name="point2">点2</param>
        /// <param name="point3">点3</param>
        /// <returns></returns>
        public AcadCircle AddCircleBy3Point1(double[] point1, double[] point2, double[] point3)
        {
            //如果三点在一条直线上无法创建圆
            if ((point1[0] - point2[0]) / (point1[1] - point2[1]) ==
                (point1[0] - point3[0]) / (point1[1] - point3[1]))
            {
                Console.WriteLine("三点共线无法创建圆");
                return null;
            }
            else
            {
                /*与圆心到三个点的距离相等，
               点1点2连线的中垂线上的点到它们的距离都相等
                因此圆心为点1点2和点1点3中垂线交点
                */
                //1.求点1点2和点1点3连线的中点
                double[] point12 = { (point1[0] + point2[0]) / 2, (point1[1] + point2[1]) / 2, 0 };
                double[] point13 = { (point1[0] + point3[0]) / 2, (point1[1] + point3[1]) / 2, 0 };
                //2.求点1点2和点1点3连线的角度，再+90度为中垂线的角度,(先画连线，再取连线的角度)
                AcadLine line12 = modelSpace.AddLine(point1, point2);
                AcadLine line13 = modelSpace.AddLine(point1, point3);
                double angle12 = line12.Angle + Math.PI / 2;
                double angle13 = line13.Angle + Math.PI / 2;
                //3.根据中点和角度创建中垂线
                AcadLine midLine12 = AddLineByReAngle(point12, angle12, 100);
                AcadLine midLine13 = AddLineByReAngle(point13, angle13, 100);
                //4.求两中垂线交点为圆心(延伸相交)
                object intersection = midLine12.IntersectWith(midLine13, AcExtendOption.acExtendBoth);
                double[] centerPoint = (double[])intersection;
                //5.点到圆心的距离就是半径
                double radius =
                    Math.Sqrt(Math.Pow(point1[0] - centerPoint[0], 2) + Math.Pow(point1[1] - centerPoint[1], 2));
                //6.清除辅助线
                line12.Delete();
                line13.Delete();
                midLine12.Delete();
                midLine13.Delete();
                //7.创建圆
                AcadCircle circle = modelSpace.AddCircle(centerPoint, radius);
                return circle;
            }
        }

        /// <summary>
        /// 通过三点创建圆，计算法
        /// </summary>
        /// <param name="point1">点1</param>
        /// <param name="point2">点2</param>
        /// <param name="point3">点3</param>
        /// <returns></returns>
        public AcadCircle AddCircleBy3Point2(double[] point1, double[] point2, double[] point3)
        {
            //参考：https://blog.csdn.net/liyuanbhu/article/details/52891868

            //如果三点在一条直线上无法创建圆，（斜率相等）
            if ((point1[0] - point2[0]) / (point1[1] - point2[1]) ==
                (point1[0] - point3[0]) / (point1[1] - point3[1]))
            {
                Console.WriteLine("三点共线无法创建圆");
                return null;
            }
            else
            {
                double x1 = point1[0], x2 = point2[0], x3 = point3[0];
                double y1 = point1[1], y2 = point2[1], y3 = point3[1];
                double a = x1 - x2;
                double b = y1 - y2;
                double c = x1 - x3;
                double d = y1 - y3;
                double e = ((x1 * x1 - x2 * x2) + (y1 * y1 - y2 * y2)) / 2.0;
                double f = ((x1 * x1 - x3 * x3) + (y1 * y1 - y3 * y3)) / 2.0;
                double det = b * c - a * d;
                if (Math.Abs(det) < 1e-5)
                {
                    Console.WriteLine("计算错误，无法创建圆");
                    return null;
                }
                double[] centerPoint = { -(d * e - b * f) / det, -(a * f - c * e) / det, 0 };
                double radius =
                    Math.Sqrt(Math.Pow(point1[0] - centerPoint[0], 2) + Math.Pow(point1[1] - centerPoint[1], 2));
                AcadCircle circle = modelSpace.AddCircle(centerPoint, radius);
                return circle;
            }
        }

        /// <summary>
        /// 绘制圆弧Demo
        /// </summary>
        /// <returns></returns>
        public AcadArc AddArcDemo()
        {
            double[] centerPoint = { 0, 0, 0 };
            double radius = 50;
            double startAngle = Math.PI / 4;
            double endAngle = 10 * Math.PI / 4; //如果终止数字大于2PI，则cad会计算终止点位置，然后逆时针方向连接起来
            //默认角度都都是以弧度计算,起始到终止为逆时针方向
            AcadArc arc = modelSpace.AddArc(centerPoint, radius, startAngle, endAngle);
            return arc;
        }

        /// <summary>
        /// 通过圆心，半径，起始角度和终止角度绘制圆弧
        /// </summary>
        /// <param name="centerPoint">圆心</param>
        /// <param name="radius">半径</param>
        /// <param name="startAngle">起始角度</param>
        /// <param name="endAngle">终止角度</param>
        /// <returns></returns>
        public AcadArc AddArcByParams(double[] centerPoint, double radius, double startAngle, double endAngle)
        {
            AcadArc arc = modelSpace.AddArc(centerPoint, radius, startAngle, endAngle);
            return arc;
        }

        /// <summary>
        /// 通过圆心，起点和端点绘制圆弧
        /// </summary>
        /// <param name="centerPoint">圆心</param>
        /// <param name="startPoint">起点</param>
        /// <param name=""></param>
        /// <param name="endPoint">端点</param>
        /// <returns></returns>
        public AcadArc AddArcByStartToEndPoint(double[] centerPoint, double[] startPoint, double[] endPoint)
        {
            double radius = Math.Sqrt(Math.Pow(centerPoint[0] - startPoint[0], 2) +
                                      Math.Pow(centerPoint[1] - startPoint[1], 2));
            double startAngle = acadDoc.Utility.AngleFromXAxis(centerPoint, startPoint); //点到点连线相对于X轴的角度（逆时针）
            double endAngle = acadDoc.Utility.AngleFromXAxis(centerPoint, endPoint);

            AcadArc arc = modelSpace.AddArc(centerPoint, radius, startAngle, endAngle);
            return arc;
        }

        /// <summary>
        /// 通过圆心，起点和角度绘制圆弧
        /// </summary>
        /// <param name="centerPoint">圆心</param>
        /// <param name="startPoint">起点</param>
        /// <param name="angle">角度</param>
        /// <returns></returns>
        public AcadArc AddArcByStartAndAngle(double[] centerPoint, double[] startPoint, double angle)
        {
            double radius = Math.Sqrt(Math.Pow(centerPoint[0] - startPoint[0], 2) +
                                      Math.Pow(centerPoint[1] - startPoint[1], 2));
            double startAngle = acadDoc.Utility.AngleFromXAxis(centerPoint, startPoint);
            double endAngle = startAngle + angle * Math.PI / 180; //换算成弧度
            AcadArc arc = modelSpace.AddArc(centerPoint, radius, startAngle, endAngle);
            return arc;
        }

        /// <summary>
        /// 通过圆心，起点和弦长绘制圆弧
        /// </summary>
        /// <param name="centerPoint">圆心</param>
        /// <param name="startPoint">起点</param>
        /// <param name="chordLength">弦长</param>
        /// <returns></returns>
        public AcadArc AddArcByStartAndChordLength(double[] centerPoint, double[] startPoint, double chordLength)
        {
            double radius = Math.Sqrt(Math.Pow(centerPoint[0] - startPoint[0], 2) +
                                      Math.Pow(centerPoint[1] - startPoint[1], 2));
            double startAngle = acadDoc.Utility.AngleFromXAxis(centerPoint, startPoint);
            double endAngle = startAngle + 2 * Math.Asin((chordLength / 2) / radius); //反正弦
            AcadArc arc = modelSpace.AddArc(centerPoint, radius, startAngle, endAngle);
            return arc;
        }

        /// <summary>
        /// 通过圆心，起点和弧长绘制圆弧
        /// </summary>
        /// <param name="centerPoint">圆心</param>
        /// <param name="startPoint">起点</param>
        /// <param name="arcLength">弧长</param>
        /// <returns></returns>
        public AcadArc AddArcByStartAndArcLength(double[] centerPoint, double[] startPoint, double arcLength)
        {
            double radius = Math.Sqrt(Math.Pow(centerPoint[0] - startPoint[0], 2) +
                                      Math.Pow(centerPoint[1] - startPoint[1], 2));
            double startAngle = acadDoc.Utility.AngleFromXAxis(centerPoint, startPoint);
            double endAngle = startAngle + arcLength / radius;
            AcadArc arc = modelSpace.AddArc(centerPoint, radius, startAngle, endAngle);
            return arc;
        }

        /// <summary>
        /// 通过三点绘制圆弧
        /// </summary>
        /// <param name="point1">点1</param>
        /// <param name="point2">点2</param>
        /// <param name="point3">点3</param>
        /// <returns></returns>
        public AcadArc AddArcBy3Point(double[] point1, double[] point2, double[] point3)
        {
            //如果三点在一条直线上无法创建圆弧
            if ((point1[0] - point2[0]) / (point1[1] - point2[1]) ==
                (point1[0] - point3[0]) / (point1[1] - point3[1]))
            {
                Console.WriteLine("三点共线无法创建圆弧");
                return null;
            }
            else
            {
                double[] point12 = { (point1[0] + point2[0]) / 2, (point1[1] + point2[1]) / 2, 0 };
                double[] point13 = { (point1[0] + point3[0]) / 2, (point1[1] + point3[1]) / 2, 0 };
                double angle12 = acadDoc.Utility.AngleFromXAxis(point1, point2) + Math.PI / 2;
                double angle13 = acadDoc.Utility.AngleFromXAxis(point1, point3) + Math.PI / 2;
                //根据中点和角度创建中垂线
                AcadLine midLine12 = AddLineByReAngle(point12, angle12, 100);
                AcadLine midLine13 = AddLineByReAngle(point13, angle13, 100);
                //求两中垂线交点为圆心(延伸相交)
                object intersection = midLine12.IntersectWith(midLine13, AcExtendOption.acExtendBoth);
                double[] centerPoint = (double[])intersection;
                //点到圆心的距离就是半径
                double radius =
                    Math.Sqrt(Math.Pow(point1[0] - centerPoint[0], 2) + Math.Pow(point1[1] - centerPoint[1], 2));
                midLine12.Delete();
                midLine13.Delete();
                double startAngle = acadDoc.Utility.AngleFromXAxis(centerPoint, point1);
                double endAngle = acadDoc.Utility.AngleFromXAxis(centerPoint, point3);
                AcadArc arc = modelSpace.AddArc(centerPoint, radius, startAngle, endAngle);
                return arc;
            }
        }

        /// <summary>
        /// 通过两点（对角线）绘制矩形
        /// </summary>
        /// <param name="point1">点1</param>
        /// <param name="point2">点2</param>
        /// <returns></returns>
        public AcadLWPolyline AddRectangleBy2Point(double[] point1, double[] point2)
        {

            //两两构成一个点,这里有三点
            double[] points = { point1[0], point1[1], point2[0], point1[1], point2[0], point2[1], point1[0], point2[1] };
            //绘制轻量多段线
            AcadLWPolyline rectangle = modelSpace.AddLightWeightPolyline(points);
            rectangle.Closed = true; //闭合
            return rectangle;
        }

        /// <summary>
        /// 通过中心点，起点和边数绘制内接多边形
        /// </summary>
        /// <param name="centerPoint">中心点</param>
        /// <param name="vertexPoint">起点</param>
        /// <param name="sideNum">边数</param>
        /// <returns></returns>
        public AcadLWPolyline AddPolygonInside(double[] centerPoint, double[] vertexPoint, int sideNum)
        {
            double radius = Math.Sqrt(Math.Pow(vertexPoint[0] - centerPoint[0], 2) +
                                      Math.Pow(vertexPoint[1] - centerPoint[1], 2)); //外接圆半径
            double angleFromX = acadDoc.Utility.AngleFromXAxis(centerPoint, vertexPoint); //顶点角度
            //多边形等分角度
            double angleDivision = 2 * Math.PI / sideNum;
            double[] points = new double[2 * sideNum];
            for (int i = 0; i < 2 * sideNum; i = i + 2)
            {
                points[i] = centerPoint[0] + Math.Cos(angleFromX) * radius;
                points[i + 1] = centerPoint[1] + Math.Sin(angleFromX) * radius;
                angleFromX = angleFromX + angleDivision;
            }
            //绘制轻量多段线
            AcadLWPolyline polygon = modelSpace.AddLightWeightPolyline(points);
            polygon.Closed = true; //闭合
            return polygon;
        }

        /// <summary>
        /// 通过中心点，相切点和边数绘制外切多边形
        /// </summary>
        /// <param name="centerPoint">中心点</param>
        /// <param name="tangentPoint">相切点</param>
        /// <param name="sideNum">边数</param>
        /// <returns></returns>
        public AcadLWPolyline AddPolygonOutside(double[] centerPoint, double[] tangentPoint, int sideNum)
        {
            double radius = Math.Sqrt(Math.Pow(tangentPoint[0] - centerPoint[0], 2) +
                                      Math.Pow(tangentPoint[1] - centerPoint[1], 2)); //内切圆半径
            double angleFromX = acadDoc.Utility.AngleFromXAxis(centerPoint, tangentPoint); //切点角度
            //多边形等分角度
            double angleDivision = 2 * Math.PI / sideNum;
            //计算顶点角度和外接圆半径
            angleFromX = angleFromX + angleDivision / 2; //顶点角度
            radius = radius / Math.Cos(angleDivision / 2); //外接圆半径
            double[] points = new double[2 * sideNum];
            for (int i = 0; i < 2 * sideNum; i = i + 2)
            {
                points[i] = centerPoint[0] + Math.Cos(angleFromX) * radius;
                points[i + 1] = centerPoint[1] + Math.Sin(angleFromX) * radius;
                angleFromX = angleFromX + angleDivision;
            }
            //绘制轻量多段线
            AcadLWPolyline polygon = modelSpace.AddLightWeightPolyline(points);
            polygon.Closed = true; //闭合
            return polygon;
        }

        /// <summary>
        /// 绘制椭圆Demo
        /// </summary>
        /// <returns></returns>
        public AcadEllipse AddEllipseDemo()
        {
            double[] centerPoint = { 10, 10, 0 };
            double[] majorAxis = { 20, 0, 0 }; //相对于中心点的坐标
            double radiusRatio = 0.5; //短轴与长轴的比值,必须小于或等于1
            if (radiusRatio > 1) radiusRatio = 1;
            AcadEllipse ellipse = modelSpace.AddEllipse(centerPoint, majorAxis, radiusRatio);
            return ellipse;
        }

        /// <summary>
        /// 通过中心点，长轴顶点和短轴长轴比绘制椭圆
        /// </summary>
        /// <param name="centerPoint">中心点</param>
        /// <param name="majorAxis">长轴顶点</param>
        /// <param name="radiusRatio">短轴长轴比</param>
        /// <returns></returns>
        public AcadEllipse AddEllipseByAxisAndRatio(double[] centerPoint, double[] majorAxis, double radiusRatio)
        {
            majorAxis[0] = majorAxis[0] - centerPoint[0]; //相对于中心点的坐标
            majorAxis[1] = majorAxis[1] - centerPoint[1];
            if (radiusRatio > 1) radiusRatio = 1;
            AcadEllipse ellipse = modelSpace.AddEllipse(centerPoint, majorAxis, radiusRatio);
            return ellipse;
        }

        /// <summary>
        /// 通过中心点，长轴顶点和短轴顶点绘制椭圆
        /// </summary>
        /// <param name="centerPoint">中心点</param>
        /// <param name="majorAxis">长轴顶点</param>
        /// <param name="vertexPoint">短轴顶点</param>
        /// <returns></returns>
        public AcadEllipse AddEllipseBy3Point(double[] centerPoint, double[] majorAxis, double[] vertexPoint)
        {
            majorAxis[0] = majorAxis[0] - centerPoint[0];
            majorAxis[1] = majorAxis[1] - centerPoint[1];
            vertexPoint[0] = vertexPoint[0] - centerPoint[0];
            vertexPoint[1] = vertexPoint[1] - centerPoint[1];
            double radiusLong = Math.Sqrt(Math.Pow(majorAxis[0], 2) + Math.Pow(majorAxis[1], 2));
            double radiusShort = Math.Sqrt(Math.Pow(vertexPoint[0], 2) + Math.Pow(vertexPoint[1], 2));
            double radiusRatio = radiusShort / radiusLong; //短轴与长轴的比值
            if (radiusRatio > 1) radiusRatio = radiusLong / radiusShort; //反过来
            AcadEllipse ellipse = modelSpace.AddEllipse(centerPoint, majorAxis, radiusRatio);
            return ellipse;
        }

        /// <summary>
        /// 绘制填充Demo
        /// </summary>
        /// <returns></returns>
        public AcadHatch AddHatchDemo()
        {
            //首先创建一个封闭区域，PolyLine
            double[] points = { 10, 10, 0, 30, 10, 0, 30, 20, 0, 10, 20, 0 };
            AcadPolyline objPolyline = modelSpace.AddPolyline(points);
            objPolyline.Closed = true;
            AcadEntity[] entArray = new AcadEntity[1];
            entArray[0] = (AcadEntity)objPolyline;
            string patternName = "ANSI33";
            AcadHatch hatch = modelSpace.AddHatch(0, patternName, true);
            hatch.AppendInnerLoop(entArray);
            //边界对象数组，边界必须包含以下类型的对象： Line, Polyline, Circle, Ellipse, Spline, Region 
            hatch.PatternScale = 0.25;
            hatch.PatternAngle = 90 * Math.PI / 180; //以弧度计算的
            hatch.color = ACAD_COLOR.acGreen;
            //当为图案填充定义定了边界后，使用 Evaluate 方法计算填充线并填充该边界，然后使用 Regen 方法更新该图案填充的显示。
            hatch.Evaluate();
            acadDoc.Regen(AcRegenType.acActiveViewport);
            return hatch;
        }

        /// <summary>
        /// 通过对象数组和填充样式名称绘制填充
        /// </summary>
        /// <param name="entArray">对象数组</param>
        /// <param name="patternName">填充样式名称</param>
        /// <returns></returns>
        public AcadHatch AddHatchByName(AcadEntity[] entArray, string patternName)
        {
            AcadHatch hatch = modelSpace.AddHatch(0, patternName, true);
            hatch.AppendInnerLoop(entArray); //对象数组
            hatch.Evaluate();
            acadDoc.Regen(AcRegenType.acActiveViewport);
            return hatch;
        }

        /// <summary>
        /// 绘制渐变填充Demo
        /// </summary>
        /// <returns></returns>
        public AcadHatch AddGradientHatchDemo()
        {
            double[] points = { 30, 10, 0, 50, 10, 0, 50, 20, 0, 30, 20, 0 };
            AcadPolyline objPolyline = modelSpace.AddPolyline(points);
            objPolyline.Closed = true;
            AcadEntity[] entArray = new AcadEntity[1];
            entArray[0] = (AcadEntity)objPolyline;
            string patternName = "LINEAR";
            AcadAcCmColor acColor1 = (AcadAcCmColor)acadDoc.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");
            AcadAcCmColor acColor2 = (AcadAcCmColor)acadDoc.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");
            acColor1.ColorIndex = (AcColor)1;
            acColor2.ColorIndex = (AcColor)2;
            AcadHatch hatch = modelSpace.AddHatch(0, patternName, true, 1); //最后一个为渐变选项
            hatch.GradientColor1 = acColor1;
            hatch.GradientColor2 = acColor2;
            hatch.AppendInnerLoop(entArray); //对象数组
            hatch.Evaluate();
            acadDoc.Regen(AcRegenType.acActiveViewport);
            return hatch;
        }

        /// <summary>
        /// 通过插入点和字高绘制单行文字
        /// </summary>
        /// <param name="textString">内容</param>
        /// <param name="insertPoint">插入点</param>
        /// <param name="height">字高</param>
        /// <returns></returns>
        public AcadText AddText(string textString, double[] insertPoint, double height)
        {
            AcadText text = modelSpace.AddText(textString, insertPoint, height);
            return text;
        }

        /// <summary>
        /// 通过插入点和宽度绘制多行文字
        /// </summary>
        /// <param name="insertPoint">插入点</param>
        /// <param name="width">宽度</param>
        /// <param name="textString">内容</param>
        /// <returns></returns>
        public AcadMText AddMText(double[] insertPoint, double width, string textString)
        {
            //默认字高从系统变量中获取
            AcadMText mText = modelSpace.AddMText(insertPoint, width, textString);
            return mText;
        }

        #endregion 绘制图元

        #region 图元属性

        /// <summary>
        /// 修改圆属性
        /// </summary>
        public void ChangeCircleDemo()
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                //防止用户非法选择导致程序崩溃
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要修改的圆：");
                    entity = (AcadEntity)objEntity;
                    if (string.Compare(entity.EntityName, "AcDbCircle") == 0)
                    {
                        entity.color = ACAD_COLOR.acBlue;
                        acadDoc.Utility.Prompt("圆修改完成。");
                    }
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
        }

        /// <summary>
        /// 修改文字
        /// </summary>
        public void ChangeTextDemo()
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要修改的文字：");
                    entity = (AcadEntity)objEntity;
                    //单行文本AcDbText；多行文本为AcDbMText
                    if (entity.EntityName == "AcDbText")
                    {
                        string str = acadDoc.Utility.GetString(0, "请输入新内容");
                        AcadText text = (AcadText)entity;
                        text.TextString = str;
                        acadDoc.Utility.Prompt("文字修改完成。");
                    }
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
        }

        /// <summary>
        /// 给现有文字添加边框
        /// </summary>
        public void AddBoundingBoxOnText()
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要添加边框的文字：");
                    entity = (AcadEntity)objEntity;
                    //单行文本AcDbText；多行文本为AcDbMText
                    if (entity.EntityName == "AcDbText")
                    {
                        AcadText text = (AcadText)entity;
                        object objMinPoint;
                        object objMaxPoint;
                        text.GetBoundingBox(out objMinPoint, out objMaxPoint);
                        AddRectangleBy2Point((double[])objMinPoint, (double[])objMaxPoint);
                        acadDoc.Utility.Prompt("边框添加完成完成。");
                    }
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
        }

        /// <summary>
        /// 通过圆心和半径绘制内接五角星
        /// </summary>
        /// <param name="centerPoint"></param>
        /// <param name="radius"></param>
        /// <returns></returns>
        public AcadLWPolyline AddStarInCircle(double[] centerPoint, double radius, out double[] points)
        {
            double angleFromX = Math.PI / 2; //起点在Y轴上
            double angleDivision = 2 * Math.PI / 5; //五角星
            points = new double[10];
            for (int i = 0; i < 10; i = i + 2)
            {
                points[i] = centerPoint[0] + Math.Cos(angleFromX) * radius;
                points[i + 1] = centerPoint[1] + Math.Sin(angleFromX) * radius;
                angleFromX = angleFromX + angleDivision * 2; //跨两个等分角（因为是绘制五角星）
            }
            //绘制轻量多段线
            AcadLWPolyline star = modelSpace.AddLightWeightPolyline(points);
            star.Closed = true; //闭合
            return star;
        }

        /// <summary>
        /// 通过选择一个现有圆绘制内接五角星
        /// </summary>
        public AcadLWPolyline AddStarInCircleDemo(out double[] points)
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            AcadLWPolyline star = null;
            points = new double[10];
            do
            {
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要绘制五角星的圆：");
                    entity = (AcadEntity)objEntity;
                    if (entity.EntityName == "AcDbCircle")
                    {
                        AcadCircle circle = (AcadCircle)entity;
                        star = AddStarInCircle((double[])circle.Center, circle.Radius, out points);
                        star.color = ACAD_COLOR.acGreen;
                        return star;
                    }
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
            return star;
        }

        /// <summary>
        /// 创建选择集
        /// </summary>
        /// <param name="selSetName">选择集名称</param>
        /// <returns></returns>
        public AcadSelectionSet CreateSelectionSet(string selSetName)
        {
            AcadSelectionSet selectionSet;
            //如果存在就先删除
            //方法一
            try
            {
                acadDoc.SelectionSets.Item(selSetName).Delete();
            }
            catch (Exception)
            {
                Console.WriteLine(selSetName + " not exist");
            }
            finally
            {
                //创建选择集
                selectionSet = acadDoc.SelectionSets.Add(selSetName);
            }
            //方法二
            //for (int i = 0; i < acadDoc.SelectionSets.Count; i++)
            //{
            //    selectionSet = acadDoc.SelectionSets.Item(i);
            //    if (string.Compare(selectionSet.Name,selSetName)==0)
            //    {
            //        selectionSet.Delete();
            //        break;
            //    }
            //}
            //selectionSet = acadDoc.SelectionSets.Add(selSetName);
            return selectionSet;
        }

        /// <summary>
        /// 选择全部图元作为选择集更改颜色
        /// </summary>
        public void SelectionSetAllEntity()
        {
            AcadSelectionSet selSet = CreateSelectionSet("mySelectionSet");
            selSet.Select(AcSelect.acSelectionSetAll); //添加全部图元，后添加的图元先加入selectionSet,堆栈的方式取图元
            AcadEntity entity;
            //遍历方法1：注意图元的出栈方式
            foreach (var item in selSet)
            {
                entity = (AcadEntity)item;
                entity.color = ACAD_COLOR.acMagenta;
            }
            //遍历方法2：
            //for (int i = 0; i < selSet.Count; i++)
            //{
            //    entity = selSet.Item(i);
            //    entity.color = ACAD_COLOR.acMagenta;
            //}
        }

        /// <summary>
        /// 通过用户选择作为选择集更改颜色
        /// </summary>
        public void SelectionSetByUser()
        {
            AcadSelectionSet selSet = CreateSelectionSet("mySelectionSet");
            //acadDoc.Utility.Prompt("请选择图元：");
            //double[] point1 = (double[])acadDoc.Utility.GetPoint();
            //double[] point2 = (double[])acadDoc.Utility.GetCorner(point1, "");
            //if (point1[0] < point2[0])
            //{
            //    selSet.Select(AcSelect.acSelectionSetWindow, point1, point2);
            //}
            //else
            //{
            //    selSet.Select(AcSelect.acSelectionSetCrossing, point1, point2);
            //}

            selSet.SelectOnScreen(); //按照用户选择的先后顺序添加图元
            AcadEntity entity;
            foreach (var item in selSet)
            {
                entity = (AcadEntity)item;
                entity.color = ACAD_COLOR.acGreen;
            }
        }

        public void SelectionSetByFilterCircle()
        {
            AcadSelectionSet selSet = CreateSelectionSet("mySelectionSet");
            //acadDoc.Utility.Prompt("请选择图元：");
            //double[] point1 = (double[])acadDoc.Utility.GetPoint();
            //double[] point2 = (double[])acadDoc.Utility.GetCorner(point1, "");
            //if (point1[0] < point2[0])
            //{
            //    selSet.Select(AcSelect.acSelectionSetWindow, point1, point2,filterType, filterData);
            //}
            //else
            //{
            //    selSet.Select(AcSelect.acSelectionSetCrossing, point1, point2, filterType, filterData);
            //}

            //还不知道怎么在C#中实现筛选!!!以后有机会补上
            //API提示指定使用的过滤器类型的 DXF 组码。 
            //CAD中命令：(setq en_data(entget(car(entsel))))
            /*选择对象: ((-1 . <图元名: -267bb0>) (0 . "LINE") (330 . <图元名: -269308>) (5 . "1C2") 
            (100. "AcDbEntity") (67. 0) (410. "Model") (8. "0") (62. 3) (100.
            "AcDbLine") (10 180.031 92.372 0.0) (11 676.345 - 182.36 0.0) (210 0.0 0.0 1.0))

            选择对象: ((-1 . <图元名: -267bc0>) (0 . "CIRCLE") (330 . <图元名: -269308>) (5 . "1C0") 
            (100 . "AcDbEntity") (67 . 0) (410 . "Model") (8 . "0") (62 . 3) (100 . 
            "AcDbCircle") (10 416.371 721.597 0.0) (40 . 75.6643) (210 0.0 0.0 1.0))
            */
            /*VBA示例代码如下：
            Dim gpCode(0) As Integer
            Dim dataValue(0) As Variant
            gpCode(0) = 0
            dataValue(0) = "Circle"

            Dim groupCode As Variant, dataCode As Variant
            groupCode = gpCode
            dataCode = dataValue

            ssetObj.Select mode, corner1, corner2, groupCode, dataCode
            */

            //int[] gpCode = new int[4];
            //string[] dataValue = new string[4];
            //gpCode[0] = -4;
            //gpCode[1] = 0;//类型
            //gpCode[2] = 8;//图层
            //gpCode[3] = -4;
            //dataValue[0] = "<or";
            //dataValue[1] = "CIRCLE";
            //dataValue[2] = "0";
            //dataValue[3] = "or>";

            int[] gpCode = new int[1];
            string[] dataValue = new string[1];

            gpCode[0] = 0;
            dataValue[0] = "CIRCLE"; //类型

            //selSet.SelectOnScreen(gpCode, dataValue);
            selSet.Select(AcSelect.acSelectionSetAll, gpCode, dataValue);

            //selSet.SelectOnScreen();
            AcadEntity entity;
            foreach (var item in selSet)
            {
                entity = (AcadEntity)item;
                entity.color = ACAD_COLOR.acMagenta;
            }
        }

        /// <summary>
        /// 手动添加图元到选择集中，过滤单行文本
        /// </summary>
        /// <returns></returns>
        public AcadSelectionSet SelectionSetBySelectText()
        {
            AcadSelectionSet selSet = CreateSelectionSet("mySelectionSet");
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                try
                {
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
                    entity = (AcadEntity)objEntity;
                    if (entity.EntityName == "AcDbText") //只要单行文本
                    {
                        AcadEntity[] entities = { entity };
                        selSet.AddItems(entities); //添加到选择集的对象数组(注意必须是数组)
                    }
                }
                catch (Exception)
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" && check != "C") flag = false;
                }
            } while (flag);
            return selSet;
        }

        /// <summary>
        /// 将选择集中的图元颜色改成红色
        /// </summary>
        public void ChangeEntityInSelSet()
        {
            AcadSelectionSet selSet = SelectionSetBySelectText();
            for (int i = 0; i < selSet.Count; i++)
            {
                AcadEntity entity = (AcadEntity)selSet.Item(i);
                entity.color = ACAD_COLOR.acRed;
            }
        }

        /// <summary>
        /// 选中文字并指定点对齐排列
        /// </summary>
        public void SortText()
        {
            AcadSelectionSet selSet = SelectionSetBySelectText();
            if (selSet.Count == 0) return;

            acadDoc.Utility.Prompt("请拾取对齐点：");
            double[] targetPoint = (double[])acadDoc.Utility.GetPoint();
            AcadText text = (AcadText)selSet.Item(0);
            text.Height = 30; //统一字高
            text.InsertionPoint = targetPoint; //文字的端点
            for (int i = 1; i < selSet.Count; i++)
            {
                text = (AcadText)selSet.Item(i);
                double[] insetPoint = (double[])text.InsertionPoint;
                insetPoint[0] = targetPoint[0];
                insetPoint[1] = targetPoint[1] - 50; //两行间距50
                targetPoint = insetPoint; //相对位置递增
                text.Height = 30;
                text.InsertionPoint = insetPoint;
            }
        }

        /// <summary>
        /// 替换文字
        /// </summary>
        /// <param name="findText">原文字</param>
        /// <param name="replaceText">要替换的文字</param>
        public void ReplaceText(string findText, string replaceText)
        {
            AcadSelectionSet selSet = CreateSelectionSet("mySelectionSet");
            //选择全部图元
            selSet.Select(AcSelect.acSelectionSetAll);
            //过滤单行文本
            if (selSet.Count == 0) return;
            AcadEntity entity;
            List<AcadEntity> entitiesList = new List<AcadEntity>();
            foreach (var item in selSet)
            {
                entity = (AcadEntity)item;
                if (entity.EntityName != "AcDbText")
                {
                    entitiesList.Add(entity);
                }
            }
            AcadEntity[] entities = entitiesList.ToArray();
            selSet.RemoveItems(entities); //排除不是文字的图元
            if (selSet.Count == 0) return;
            for (int i = 0; i < selSet.Count; i++)
            {
                AcadText text = (AcadText)selSet.Item(i);
                string oldText = text.TextString; //获取文本
                if (oldText.Contains(findText))
                {
                    string newText = oldText.Replace(findText, replaceText); //替换文字
                    Debug.Print(newText);
                    text.TextString = newText;
                }
            }
        }

        /// <summary>
        /// 统计图元个数
        /// </summary>
        public void EntitiesStatistics()
        {
            AcadSelectionSet selSet = CreateSelectionSet("mySelectionSet");
            //选择全部图元
            selSet.Select(AcSelect.acSelectionSetAll);
            if (selSet.Count == 0) return;
            Dictionary<string, int> entDic =
                new Dictionary<string, int>();
            foreach (var item in selSet)
            {
                AcadEntity entity = (AcadEntity)item;
                if (entDic.ContainsKey(entity.EntityName)) entDic[entity.EntityName] += 1;
                else entDic.Add(entity.EntityName, 1);

            }
            foreach (var item in entDic)
            {
                Console.WriteLine("图元类型 " + item.Key + " : " + item.Value + " 个");
            }
        }

        /// <summary>
        /// 根据用户选择对象，统计材料小插件
        /// </summary>
        public void MaterialStatistics()
        {
            //手动对象到选择集
            AcadSelectionSet selSet = CreateSelectionSet("mySelectionSet");
            selSet.SelectOnScreen();
            //过滤单行文本
            if (selSet.Count == 0) return;
            List<AcadEntity> entitiesList = new List<AcadEntity>();
            foreach (var item in selSet)
            {
                AcadEntity entity = (AcadEntity)item;
                if (entity.EntityName != "AcDbText")
                {
                    entitiesList.Add(entity);
                }
            }
            AcadEntity[] entities = entitiesList.ToArray();
            selSet.RemoveItems(entities); //排除不是文字的图元
            //固定的几种材料
            Dictionary<string, int> matDic =
                new Dictionary<string, int>();
            foreach (var item in selSet)
            {
                AcadText text = (AcadText)item;
                if (matDic.ContainsKey(text.TextString)) matDic[text.TextString] += 1;
                else matDic.Add(text.TextString, 1);
            }
            //输出显示
            foreach (var item in matDic)
            {
                //判断这个文本是不是材料关键字SS或GI
                if (item.Key.Contains("SS") || item.Key.Contains("GI"))
                {
                    Console.WriteLine("材料 " + item.Key + " : " + item.Value + " PCS");
                }
            }
        }

        #endregion 图元属性

        #region 图元修改

        /// <summary>
        /// 移动实体Demo
        /// </summary>
        public void MoveEntityDemo()
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要移动的实体：");
                    entity = (AcadEntity)objEntity;
                    double[] startPoint = { 0, 0, 0 };
                    double[] endPoint = { 100, 0, 0 };
                    entity.Move(startPoint, endPoint); //移动实体
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
        }

        /// <summary>
        /// 旋转实体Demo
        /// </summary>
        public void RotateEntityDemo()
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要旋转的实体：");
                    entity = (AcadEntity)objEntity;
                    double[] basePoint = { 0, 0, 0 };
                    double rotateAngle = 45 * Math.PI / 180; //给定的是弧度
                    entity.Rotate(basePoint, rotateAngle); //旋转实体
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
        }

        /// <summary>
        /// 删除实体Demo
        /// </summary>
        public void DeleteEntityDemo()
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要删除的实体：");
                    entity = (AcadEntity)objEntity;
                    entity.Delete(); //删除实体
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
        }

        /// <summary>
        /// 调用CAD命令，不足：只管发命令，不管执行结果
        /// </summary>
        /// <param name="strCommand"></param>
        public void SendCommandToCAD(string strCommand)
        {
            acadDoc.SendCommand(strCommand);
        }

        /// <summary>
        /// 调用CAD命令绘制直线，圆，文本
        /// </summary>
        public void SendCommandToCADDemo()
        {
            SendCommandToCAD("line\r0,0,0\r100,100,0\r\r"); //\r代表回车
            SendCommandToCAD("circle\r100,100,0\r50\r");
        }

        /// <summary>
        /// 调用CAD命令绘制直线，用户选择起点和终点
        /// </summary>
        public void SendCommandToCADAddLineByUser()
        {
            acadDoc.Utility.Prompt("请选择起点：");
            double[] startPoint = (double[])acadDoc.Utility.GetPoint();
            double[] endPoint = (double[])acadDoc.Utility.GetPoint(startPoint, "请选择终点：");
            string strCommand = "line\r" + ChangePointFormat(startPoint) + "\r" + ChangePointFormat(endPoint) + "\r\r";
            SendCommandToCAD(strCommand);
        }

        /// <summary>
        /// 将点变换成字符串格式（lisp点，坐标）
        /// </summary>
        /// <param name="point">点</param>
        /// <returns></returns>
        public string ChangePointFormat(double[] point)
        {
            return point[0] + "," + point[1] + "," + point[2];
        }

        /* 命令: (setq en(entsel))
         * 选择对象: (<图元名: -2677c8> (312.735 1054.97 0.0))
         * (list (handent "276") (list 100 100 0))
         */
        /// <summary>
        /// 将实体换换成字符串格式（lisp图元 对象）
        /// </summary>
        /// <param name="entity">对象</param>
        /// <returns></returns>
        public string ChangeEntityFormat(AcadEntity entity)
        {
            return "(handent \"" + entity.Handle + "\")";
        }

        /// <summary>
        /// 将实体换换成字符串格式（lisp图元 对象，带坐标）
        /// </summary>
        /// <param name="entity">对象</param>
        /// <returns></returns>
        public string ChangeEntityAndPointFormat(AcadEntity entity, double[] point)
        {
            return "(list (handent \"" + entity.Handle + "\") (list " + point[0] + " " + point[1] + " " + point[2] +
                   "))";
        }

        /// <summary>
        /// 调用CAD命令，打断命令
        /// </summary>
        public void SendCommandToCADBreak()
        {
            object objEntity;
            object pickedPoint;
            AcadEntity entity;
            Boolean flag = true;
            do
            {
                try
                {
                    //让用户选取对象
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要打断的实体：");
                    entity = (AcadEntity)objEntity;
                    double[] startPoint = (double[])pickedPoint;
                    double[] endPoint = (double[])acadDoc.Utility.GetPoint(startPoint, "请选择打断第二点：");
                    string strCommand = "break\r";
                    strCommand += ChangeEntityAndPointFormat(entity, startPoint) + "\r";
                    strCommand += ChangePointFormat(endPoint) + "\r";
                    SendCommandToCAD(strCommand);
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C继续，任意键结束：");
                    if (check != "c" || check != "C") flag = false;
                }
            } while (flag);
        }

        /// <summary>
        /// 调用CAD命令，修剪命令，修剪五角星
        /// </summary>
        public void TrimStar()
        {
            double[] centerPoint = { 0, 0, 0 };
            double[] points = new double[10];
            double angleFromX = Math.PI / 2; //起点在Y轴上
            double angleDivision = 2 * Math.PI / 5; //五角星
            for (int i = 0; i < 10; i = i + 2)
            {
                points[i] = centerPoint[0] + Math.Cos(angleFromX) * 50;
                points[i + 1] = centerPoint[1] + Math.Sin(angleFromX) * 50;
                angleFromX = angleFromX + angleDivision * 2; //跨两个等分角（因为是绘制五角星）
            }
            //绘制轻量多段线
            AcadLWPolyline star = modelSpace.AddLightWeightPolyline(points);
            star.Closed = true; //闭合
            AcadEntity entity = (AcadEntity)star;
            string strCommand = "trim\r"; //修剪命令
            strCommand += ChangeEntityFormat(entity) + "\r\r";
            for (int i = 0; i < points.Length; i = i + 2)
            {
                strCommand += MidPointFormat(points[i], points[i + 1], points[(i + 2) % 10], points[(i + 3) % 10]) +
                              "\r";
            }
            strCommand += "\r";
            SendCommandToCAD(strCommand);
        }

        /// <summary>
        /// 将点变换成字符串格式，两点的中点（坐标）
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        /// <returns></returns>
        public string MidPointFormat(double x1, double y1, double x2, double y2)
        {
            return (x1 + x2) / 2 + "," + (y1 + y2) / 2 + ",0";
        }

        /// <summary>
        /// 调用CAD命令，延伸命令
        /// </summary>
        public void SendCommandToCADExtend()
        {
            Object objEntity;
            Object pickedPoint;
            AcadEntity entity;
            acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
            entity = (AcadEntity)objEntity;
            string strCommand = "extend\r\r";
            //strCommand += ChangeEntityFormat(entity) + ChangePointFormat((double[])pickedPoint) + "\r";
            strCommand += ChangePointFormat((double[])pickedPoint) + "\r";
            SendCommandToCAD(strCommand);
        }

        /// <summary>
        /// 复制实体
        /// </summary>
        public void CopyEntityDemo()
        {
            Object objEntity;
            Object pickedPoint;
            AcadEntity entity;
            acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
            entity = (AcadEntity)objEntity;
            AcadEntity copyEntity = (AcadEntity)entity.Copy();
            acadDoc.Utility.Prompt("请选择复制起点：");
            double[] startPoint = (double[])acadDoc.Utility.GetPoint();
            double[] endPoint = (double[])acadDoc.Utility.GetPoint(startPoint, "请选择复制终点：");
            copyEntity.Move(startPoint, endPoint);
            copyEntity.color = ACAD_COLOR.acRed;
        }

        /// <summary>
        /// 镜像实体
        /// </summary>
        public void MirrorEntityDemo()
        {
            Object objEntity;
            Object pickedPoint;
            AcadEntity entity;
            acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
            entity = (AcadEntity)objEntity;
            acadDoc.Utility.Prompt("请选择镜像中心线起点：");
            double[] startPoint = (double[])acadDoc.Utility.GetPoint();
            double[] endPoint = (double[])acadDoc.Utility.GetPoint(startPoint, "请选择镜像中心线终点：");
            AcadEntity mirrorEntity = (AcadEntity)entity.Mirror(startPoint, endPoint);
            mirrorEntity.color = ACAD_COLOR.acGreen;
        }

        /// <summary>
        /// 缩放实体
        /// </summary>
        public void ScaleEntityDemo()
        {
            Object objEntity;
            Object pickedPoint;
            AcadEntity entity;
            acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
            entity = (AcadEntity)objEntity;
            acadDoc.Utility.Prompt("请选择基点：");
            double[] basePoint = (double[])acadDoc.Utility.GetPoint();
            string scale = acadDoc.Utility.GetString(0, "请输入放大倍数");
            entity.ScaleEntity(basePoint, Convert.ToDouble(scale));
        }

        /// <summary>
        /// 矩形阵列
        /// </summary>
        public void ArrayRectangularDemo()
        {
            Object objEntity;
            Object pickedPoint;
            AcadEntity entity;
            acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
            entity = (AcadEntity)objEntity;
            int rows = 5;
            int columns = 10;
            double disRows = 200;
            double disColumns = 200;
            entity.ArrayRectangular(rows, columns, 1, disRows, disColumns, 1);
        }

        /// <summary>
        /// 环形阵列
        /// </summary>
        public void ArrayPolarDemo()
        {
            Object objEntity;
            Object pickedPoint;
            AcadEntity entity;
            acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择图元：");
            entity = (AcadEntity)objEntity;
            acadDoc.Utility.Prompt("请选择圆形阵列中心点点：");
            double[] centerPoint = (double[])acadDoc.Utility.GetPoint();
            int num = 10;
            double angle = 2 * Math.PI; //弧度
            entity.ArrayPolar(num, angle, centerPoint);
        }

        /// <summary>
        /// 连接合并多段线
        /// </summary>
        public void JoinLWPolyline()
        {
            AcadSelectionSet selSet = CreateSelectionSet("MySelSet");
            selSet.SelectOnScreen();
            string strCommand = "pedit\r";
            strCommand += "m\r" + ChangeSelSetFormat(selSet) + "\rj\r\r\r";
            SendCommandToCAD(strCommand);
        }

        /// <summary>
        /// 将选择集转换成字符串格式（lisp选择集分散成单个图元对象）
        /// </summary>
        /// <param name="selSet">选择集</param>
        /// <returns></returns>
        public string ChangeSelSetFormat(AcadSelectionSet selSet)
        {
            if (selSet.Count == 0) return null;
            string strSelSet = "";
            for (int i = 0; i < selSet.Count; i++)
            {
                strSelSet += "(handent \"" + selSet.Item(i).Handle + "\")\r";
            }
            return strSelSet;
        }

        /// <summary>
        /// 通过拾取封闭区域中的点，创建封闭区域
        /// </summary>
        /// <returns></returns>
        public AcadEntity CreateBoundary()
        {
            acadDoc.Utility.Prompt("请选拾取区域内部点：");
            double[] innerPoint = (double[])acadDoc.Utility.GetPoint();
            string strCommand = "-boundary\r";
            strCommand += innerPoint[0] + "," + innerPoint[1] + "\r\r"; //这个命令只需要二维点坐标
            SendCommandToCAD(strCommand);
            int entCount = modelSpace.Count; //统计图元个数
            AcadEntity entity = modelSpace.Item(entCount - 1); //获取最新创建的对象
            entity.color = ACAD_COLOR.acMagenta;
            return entity;
        }

        /* Bulge凸度：指多段线的一个顶点和下一个顶点形成的圆弧的角度的1/4的正切值
         */
        public void GetBulgeInLWPolylineDemo()
        {
            Object objEntity;
            Object pickedPoint;
            AcadEntity entity;

            acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择多段线：");
            entity = (AcadEntity)objEntity;
            int pointNum = GetPolylinePointNum(entity);
            double[] bulge = new double[pointNum];
            for (int i = 0; i < pointNum; i++)
            {
                if (entity.EntityName == "AcDbPolyline")
                    bulge[i] = ((AcadLWPolyline)entity).GetBulge(i);
                //if (entity.EntityName == "AcDbPolyline")
                //    bulge[i] = ((AcadPolyline)entity).GetBulge(i);
                Console.WriteLine(bulge[i]);
            }
        }
        /// <summary>
        /// 获取多段线的顶点数
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public int GetPolylinePointNum(AcadEntity entity)
        {
            int pointNum = 0;
            double[] points;
            if (entity.EntityName == "AcDbPolyline")
            {
                points = (double[])((AcadLWPolyline)entity).Coordinates;
                pointNum = points.Length / 2;
            }
            //if (entity.EntityName == "AcDbPolyline")
            //{
            //    points = (double[])((AcadPolyline)entity).Coordinates;
            //    pointNum = points.Length / 3;
            //}
            return pointNum;
        }
        /// <summary>
        /// 通过控制凸度绘制圆弧多段线
        /// </summary>
        /// <returns></returns>
        public AcadLWPolyline AddPolylineByBulgeDemo()
        {
            double[] points = { 0, 0, 1000, 0, 2000, 0, };
            AcadLWPolyline lwPolyline = modelSpace.AddLightWeightPolyline(points);
            lwPolyline.SetBulge(0, -0.5);//逆时针设置凸度
            lwPolyline.SetBulge(1, 0.5);//设置成1为一个半圆
            return lwPolyline;
        }
        /// <summary>
        /// 将直线转化成多段线,获取直线属性然后创建多段线，最后删除直线
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public AcadLWPolyline LineToLWPolyline(AcadLine line)
        {
            double[] startPoint = (double[])line.StartPoint;
            double[] endPoint = (double[])line.EndPoint;
            double[] points = { startPoint[0], startPoint[1], endPoint[0], endPoint[0] };
            AcadLWPolyline lwPolyline = modelSpace.AddLightWeightPolyline(points);
            line.Delete();
            return lwPolyline;
        }
        /// <summary>
        /// 将圆转化为多段线，获取圆的属性，然后创建多段线，最后删除圆
        /// </summary>
        /// <param name="circle"></param>
        /// <returns></returns>
        public AcadLWPolyline CircleToLWPolyline(AcadCircle circle)
        {
            double[] centerPoint = (double[])circle.Center;
            double[] points =
                {centerPoint[0] - circle.Radius, centerPoint[1], centerPoint[0] + circle.Radius, centerPoint[1]};
            AcadLWPolyline lwPolyline = modelSpace.AddLightWeightPolyline(points);
            lwPolyline.Closed = true;//设置闭合
            lwPolyline.SetBulge(0, 1);
            lwPolyline.SetBulge(1, 1);
            circle.Delete();
            return lwPolyline;
        }
        /// <summary>
        /// 将圆弧转化为多段线，获取圆弧的属性，然后创建多段线，最后删除圆弧
        /// </summary>
        /// <param name="arc"></param>
        /// <returns></returns>
        public AcadLWPolyline ArcToLWPolyline(AcadArc arc)
        {
            double[] centerPoint = (double[])arc.Center;

            double[] points =
                {centerPoint[0] + arc.Radius*Math.Cos(arc.StartAngle), centerPoint[1]+ arc.Radius*Math.Sin(arc.StartAngle), centerPoint[0] + + arc.Radius*Math.Cos(arc.EndAngle), centerPoint[1]+ arc.Radius*Math.Sin(arc.EndAngle)};
            AcadLWPolyline lwPolyline = modelSpace.AddLightWeightPolyline(points);
            lwPolyline.SetBulge(0, Math.Tan(arc.TotalAngle / 4));//圆弧的角度的1/4的正切值
            arc.Delete();
            return lwPolyline;
        }

        #endregion 图元修改

        #region 图层设置
        /// <summary>
        /// 添加图层
        /// </summary>
        /// <param name="str">名称</param>
        /// <returns></returns>
        public AcadLayer AddLayer(string str)
        {
            return acadDoc.Layers.Add(str);
        }
        /// <summary>
        /// 删除图层
        /// </summary>
        /// <param name="str"></param>
        public void DeleteLayer(string str)
        {
            acadDoc.Layers.Item(str).Delete();
        }
        /// <summary>
        /// 图层更名
        /// </summary>
        /// <param name="oldName"></param>
        /// <param name="newName"></param>
        public void RenameLayer(string oldName, string newName)
        {
            acadDoc.Layers.Item(oldName).Name = newName;
        }
        /// <summary>
        /// 图层设置示例
        /// </summary>
        public void LayerSettingDemo()
        {
            for (int i = 0; i < acadDoc.Layers.Count; i++)
            {
                if (acadDoc.Layers.Item(i).Name == "中心线")
                {
                    AcadLayer layer = acadDoc.Layers.Item(i);
                    layer.LayerOn = true;
                    layer.Freeze = false;
                    layer.Lock = false;
                    layer.color = AcColor.acRed;//颜色
                    layer.Lineweight = ACAD_LWEIGHT.acLnWt050;//线型

                    Boolean flag = true;
                    for (int j = 0; j < acadDoc.Linetypes.Count; j++)
                    {
                        if (acadDoc.Linetypes.Item(j).Name == "CENTER") flag = false;
                    }
                    if (flag) acadDoc.Linetypes.Load("CENTER", "acad.lin");//先加载线型再设置线型

                    layer.Linetype = "CENTER";//设置线型
                    acadDoc.ActiveLayer = layer;//设置为当前图层
                }
            }
        }
        /// <summary>
        /// 批量删除空白图层
        /// </summary>
        public void DeleteBlankLayer()
        {
            //不能删除当前图层，将其他图层设置激活图层再删除。
            //不能删除有内容的图层，遍历删除图层中的图元，然后删除图层。
            acadDoc.ActiveLayer = acadDoc.Layers.Item(0);//设置0层为当前图层
            int layerNum = acadDoc.Layers.Count;
            //现将所有图层对象装入数组中，防止直接从layers集合中删除导致循环失败
            AcadLayer[] layerArray = new AcadLayer[layerNum - 1];
            for (int i = 0; i < layerNum - 1; i++)
            {
                layerArray[i] = acadDoc.Layers.Item(i + 1);
            }
            AcadSelectionSet selSet = CreateSelectionSet("MySelSet");
            selSet.Select(AcSelect.acSelectionSetAll);
            //int[] filterType = new int[1];
            //string[] FilterData = new string[1];
            //filterType[0] = 8;
            for (int i = 0; i < layerArray.Length; i++)
            {
                //方法一：曲线救国
                int entityNum = 0;
                for (int j = 0; j < selSet.Count; j++)
                {
                    AcadEntity entity = selSet.Item(j);
                    if (entity.Layer == layerArray[i].Name) entityNum++;//计算图层中图元的个数
                }
                if (entityNum == 0) layerArray[i].Delete();//图元个数为0则删除图层

                //方法二，使用选择集过滤（再C#中COM报错，暂未实现过滤）
                //selSet.Select(AcSelect.acSelectionSetAll,null,null, filterType, FilterData);//DXF过滤设置为图层
                //if (selSet.Count==0) layerArray[i].Delete();
                //selSet.Clear();//每次循环后清除选择集
            }
        }
        #endregion 图层设置

        #region 图块操作

        /// <summary>
        /// 创建块
        /// </summary>
        public void CreateBlockDemo()
        {
            DeleteBlock("块实例");
            double[] basePoint = { 0, 0, 0 };//插入块的基点
            AcadBlock objBlock = acadDoc.Blocks.Add(basePoint, "块实例");
            double[] point1 = { -20, 0, 0 };
            double[] point2 = { 20, 0, 0 };
            double[] point3 = { 0, -20, 0 };
            double[] point4 = { 0, 20, 0 };

            objBlock.AddLine(point1, point2);
            objBlock.AddLine(point3, point4);
            objBlock.AddCircle(basePoint, 20);
            objBlock.Item(2).Lineweight = ACAD_LWEIGHT.acLnWt030;
        }
        /// <summary>
        /// 选择图元并创建块
        /// </summary>
        public void CreateBlockBySelect()
        {
            DeleteBlock("选择后创建块");
            acadDoc.Utility.Prompt("请拾取基点：");
            double[] basePoint = (double[])acadDoc.Utility.GetPoint();
            AcadBlock objBlock = acadDoc.Blocks.Add(basePoint, "选择后创建块");
            //选择要加入块的图元
            AcadSelectionSet selSet = CreateSelectionSet("MySelSet");
            selSet.SelectOnScreen(); //按照用户选择的先后顺序添加图元
            AcadEntity[] entitys = new AcadEntity[selSet.Count];
            for (int i = 0; i < selSet.Count; i++)
            {
                entitys[i] = selSet.Item(i);
            }
            acadDoc.CopyObjects(entitys, objBlock);//通过复制实体的方式插入块
            selSet.Clear();
        }
        /// <summary>
        /// 通过块名和插入点插入块
        /// </summary>
        /// <param name="blockName"></param>
        /// <param name="insertPoint"></param>
        public AcadBlockReference InsertBlock(string blockName, double[] insertPoint)
        {
            return modelSpace.InsertBlock(insertPoint, blockName, 1, 1, 1, 0);
        }
        /// <summary>
        /// 删除块
        /// </summary>
        public void DeleteBlock(string blockName)
        {
            //不能删除已经被插入的图块（对象被参照），必须先删除模型控件中的参照
            //判断是否存在该块
            Boolean flag = false;
            for (int i = 0; i < acadDoc.Blocks.Count; i++)
            {
                AcadBlock item = acadDoc.Blocks.Item(i);
                if (item.Name == blockName) flag = true;
            }
            if (flag)
            {
                //判断模型空间中是否存在块对参照，如果有则删除它
                //AcDbBlockReference

                AcadBlock objBlock = acadDoc.Blocks.Item(blockName);
                objBlock.Delete();
            }
        }

        /// <summary>
        /// 创建带属性的块
        /// </summary>
        public void CreateAttBlockDemo()
        {
            //如果不删除，会在原有的块上重复绘制
            DeleteBlock("带属性的块");
            double[] basePoint = { 0, 0, 0 };//插入块的基点
            AcadBlock objBlock = acadDoc.Blocks.Add(basePoint, "带属性的块");
            double[] point1 = { -20, 0, 0 };
            double[] point2 = { 20, 0, 0 };
            double[] point3 = { 0, -20, 0 };
            double[] point4 = { 0, 20, 0 };

            objBlock.AddLine(point1, point2);
            objBlock.AddLine(point3, point4);
            objBlock.AddCircle(basePoint, 20);
            objBlock.Item(2).Lineweight = ACAD_LWEIGHT.acLnWt030;
            //设置块的属性
            double[] basePoint1 = { 25, 2, 0 };
            objBlock.AddAttribute(16, AcAttributeMode.acAttributeModeLockPosition, "ANT1-1F", basePoint1, "编号", "ANT1-1F");
            double[] basePoint2 = { 25, -18, 0 };
            objBlock.AddAttribute(16, AcAttributeMode.acAttributeModeLockPosition, "10.0dBm", basePoint2, "功率", "10.0dBm");

        }
        /// <summary>
        /// 批量修改带属性的块
        /// </summary>
        public void ChangeBlockAtt()
        {
            //先选择需要修改的块
            AcadSelectionSet selSet = CreateSelectionSet("MySelSet");
            selSet.SelectOnScreen();
            if (selSet.Count == 0) return;

            acadDoc.Utility.Prompt("请拾取对齐点：");
            double[] targetPoint = (double[])acadDoc.Utility.GetPoint();
            for (int i = 0; i < selSet.Count; i++)
            {
                AcadBlockReference block = (AcadBlockReference)selSet.Item(i);
                //排序与文字排序一样
                double[] insetPoint = (double[])block.InsertionPoint;
                insetPoint[0] = targetPoint[0];
                insetPoint[1] = targetPoint[1] - 60; //两行间距60
                targetPoint = insetPoint; //相对位置递增
                block.InsertionPoint = targetPoint;
                //修改属性C#无法实现
            }
        }




        #endregion 图块操作




    }
}