using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Data;

namespace AcadStudy
{
    /*《ActiveX 和 VBA 参考》由明经通道翻译并提供：https://blog.csdn.net/weixin_46656590/article/details/105247189 
    下载地址：https://yunpan.360.cn/surl_yRE3CKKEqLk （提取码：c16e）
    BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
    代码由VBA的逻辑转化而来
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
            Object pointPicked;
            AcadEntity objAcadEntity;
            do
            {
                try
                {
                    acadDoc.Utility.GetEntity(out objEntity, out pointPicked, "请选择图元：");
                    objAcadEntity = (AcadEntity)objEntity;
                    Console.WriteLine(objAcadEntity.EntityName);
                }
                catch (Exception e)
                {
                    string check = acadDoc.Utility.GetString(0, "输入C结束，任意键继续：");
                    if (check == "c" || check == "C") return;
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
            double radius = Math.Sqrt(Math.Pow(vertexPoint[0] - centerPoint[0], 2) + Math.Pow(vertexPoint[1] - centerPoint[1], 2));//外接圆半径
            double angleFromX = acadDoc.Utility.AngleFromXAxis(centerPoint, vertexPoint);//顶点角度
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
            double radius = Math.Sqrt(Math.Pow(tangentPoint[0] - centerPoint[0], 2) + Math.Pow(tangentPoint[1] - centerPoint[1], 2));//内切圆半径
            double angleFromX = acadDoc.Utility.AngleFromXAxis(centerPoint, tangentPoint);//切点角度
            //多边形等分角度
            double angleDivision = 2 * Math.PI / sideNum;
            //计算顶点角度和外接圆半径
            angleFromX = angleFromX + angleDivision / 2;//顶点角度
            radius = radius / Math.Cos(angleDivision / 2);//外接圆半径
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
            double[] majorAxis = { 20, 0, 0 };//相对于中心点的坐标
            double radiusRatio = 0.5;//短轴与长轴的比值,必须小于或等于1
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
            majorAxis[0] = majorAxis[0] - centerPoint[0];//相对于中心点的坐标
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
            double radiusRatio = radiusShort / radiusLong;//短轴与长轴的比值
            if (radiusRatio > 1) radiusRatio = radiusLong / radiusShort;//反过来
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
            hatch.PatternAngle = 90 * Math.PI / 180;//以弧度计算的
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
            hatch.AppendInnerLoop(entArray);//对象数组
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
            AcadHatch hatch = modelSpace.AddHatch(0, patternName, true, 1);//最后一个为渐变选项
            hatch.GradientColor1 = acColor1;
            hatch.GradientColor2 = acColor2;
            hatch.AppendInnerLoop(entArray);//对象数组
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
                    string check = acadDoc.Utility.GetString(0, "输入C结束，任意键继续：");
                    if (check == "c" || check == "C") flag = false;
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
                    string check = acadDoc.Utility.GetString(0, "输入C结束，任意键继续：");
                    if (check == "c" || check == "C") flag = false;
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
                    entity = (AcadEntity) objEntity;
                    //单行文本AcDbText；多行文本为AcDbMText
                    if (entity.EntityName == "AcDbText")
                    {
                        AcadText text = (AcadText) entity;
                        object objMinPoint;
                        object objMaxPoint;
                        text.GetBoundingBox(out objMinPoint, out objMaxPoint);
                        AddRectangleBy2Point((double[]) objMinPoint, (double[]) objMaxPoint);
                        acadDoc.Utility.Prompt("边框添加完成完成。");
                    }
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C结束，任意键继续：");
                    if (check == "c" || check == "C") flag = false;
                }
            } while (flag);
        }
        /// <summary>
        /// 通过圆心和半径绘制内接五角星
        /// </summary>
        /// <param name="centerPoint"></param>
        /// <param name="radius"></param>
        /// <returns></returns>
        public AcadLWPolyline AddStarInCircle(double[] centerPoint, double radius)
        {
            double angleFromX = Math.PI / 2;//起点在Y轴上
            double angleDivision = 2 * Math.PI / 5;//五角星
            double[] points = new double[10];
            for (int i = 0; i < 10; i = i + 2)
            {
                points[i] = centerPoint[0] + Math.Cos(angleFromX) * radius;
                points[i + 1] = centerPoint[1] + Math.Sin(angleFromX) * radius;
                angleFromX = angleFromX + angleDivision * 2;//跨两个等分角（因为是绘制五角星）
            }
            //绘制轻量多段线
            AcadLWPolyline star = modelSpace.AddLightWeightPolyline(points);
            star.Closed = true; //闭合
            return star;
        }
        /// <summary>
        /// 通过选择一个现有圆绘制内接五角星
        /// </summary>
        public void AddStarInCircleDemo()
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
                    acadDoc.Utility.GetEntity(out objEntity, out pickedPoint, "请选择要绘制五角星的圆：");
                    entity = (AcadEntity)objEntity;
                    if (entity.EntityName == "AcDbCircle")
                    {
                        AcadCircle circle = (AcadCircle)entity;
                       AcadLWPolyline star= AddStarInCircle((double[])circle.Center, circle.Radius);
                        star.color = ACAD_COLOR.acGreen;
                    }
                }
                catch (Exception)
                {
                    acadDoc.Utility.Prompt("未选中实体，");
                }
                finally
                {
                    string check = acadDoc.Utility.GetString(0, "输入C结束，任意键继续：");
                    if (check == "c" || check == "C") flag = false;
                }
            } while (flag);
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
            //try
            //{
            //    acadDoc.SelectionSets.Item(selSetName).Delete();
            //}
            //catch (Exception)
            //{
            //    Console.WriteLine(selSetName+" not exist");
            //}
            //finally
            //{
            //    //创建选择集
            //    selectionSet = acadDoc.SelectionSets.Add(selSetName);
            //}
            //方法二
            for (int i = 0; i < acadDoc.SelectionSets.Count; i++)
            {
                selectionSet = acadDoc.SelectionSets.Item(i);
                if (string.Compare(selectionSet.Name,selSetName)==0)
                {
                    selectionSet.Delete();
                    break;
                }
            }
            selectionSet = acadDoc.SelectionSets.Add(selSetName);
            return selectionSet;
        }
        /// <summary>
        /// 选择集选择全部图元改颜色Demo
        /// </summary>
        public void SelectionSetDemo()
        {
            AcadSelectionSet selSet = CreateSelectionSet("SelectionSet");
            selSet.Select(AcSelect.acSelectionSetAll);//添加全部图元
            AcadEntity entity;
            foreach (var item in selSet)
            {
                entity = (AcadEntity)item;
                entity.color = ACAD_COLOR.acMagenta;
            }
        }





        #endregion 图元属性




    }
}
