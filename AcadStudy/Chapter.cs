using Autodesk.AutoCAD.Interop.Common;

namespace AcadStudy
{
    //BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
    //------测试代码---------------------
    public class Chapter
    {
        private AcadHelper Acad = new AcadHelper();
        /// <summary>
        /// 第1-4节课,创建直线
        /// </summary>
        public void P1to4()
        {
            double[] startPoint2 = { 50, 50, 0 };
            double[] endPoint2 = { -50, 50, 0 };
            AcadLine line1 = Acad.AddLineDemo();
            AcadLine line2 = Acad.AddLineByPoint(startPoint2, endPoint2);
            Acad.AddLineByXY(-50, 50, 0, 0);
            Acad.AddLineByReXY(startPoint2, 0, 100);
            Acad.AddLineByReAngle(endPoint2, 45, 100);
            line1.Lineweight = ACAD_LWEIGHT.acLnWt030;
            line2.color = ACAD_COLOR.acRed;
            //demo.ShowEntity();
            Acad.Zoom();
        }
        /// <summary>
        /// 第5节课，创建多段线
        /// </summary>
        public void P5()
        {
            Acad.AddLWPolylineDemo();
            Acad.AddPolylineDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第6节课，创建圆
        /// </summary>
        public void P6()
        {
            double[] centerPoint = { 0, 0, 0 };
            double[] point1 = { 50, 0, 0 };
            double[] point2 = { 150, 0, 0 };
            Acad.AddCircleDemo();
            AcadCircle railCircle = Acad.AddCircleByRadius(centerPoint, 100);
            Acad.AddCircleByDiameter(centerPoint, 300);
            AcadCircle moveCircle = Acad.AddCircleBy2Point(point1, point2);
            railCircle.color = ACAD_COLOR.acRed;
            moveCircle.color = ACAD_COLOR.acGreen;
            Acad.Zoom();
        }
        /// <summary>
        /// 第7节课，通过三点创建圆
        /// </summary>
        public void P7()
        {
            double[] point1 = { 50, 0, 0 };
            double[] point2 = { 150, 150, 0 };
            double[] point3 = { 50, 50, 0 };
            Acad.AddCircleBy3Point1(point1, point2, point3);
            AcadCircle circle = Acad.AddCircleBy3Point2(point1, point2, point3);
            circle.color = ACAD_COLOR.acGreen;
            circle.Move(point1, point2);//移开这个圆方便观察
            Acad.Zoom();
        }
        /// <summary>
        /// 第8节课，使用循环方法创建同心圆
        /// </summary>
        public void P8_1()
        {
            //手动绘制效率不及程序
            double[] centerPoint = { 0, 0, 0 };
            double radius = 50;
            for (int i = 0; i < 999; i++)
            {
                Acad.AddCircleByRadius(centerPoint, radius + i * 10);
            }
            Acad.Zoom();
        }
        /// <summary>
        /// 第8节课，使用循环方法创建阵列
        /// </summary>
        public void P8_2()
        {
            double[] centerPoint = { 0, 0, 0 };
            double radius = 50;
            for (int i = 0; i < 20; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    centerPoint[0] = i * 100;
                    centerPoint[1] = j * 100;
                    Acad.AddCircleByRadius(centerPoint, radius);
                }
            }
            Acad.Zoom();
        }
        /// <summary>
        /// 第9-10节课，创建圆弧
        /// </summary>
        public void P9to10()
        {
            double[] centerPoint = { 10, 0, 0 };
            double[] startPoint = { 50, 50, 0 };
            double[] endPoint = { 10, 10, 0 };
            Acad.AddArcDemo();
            Acad.AddArcByStartToEndPoint(centerPoint, startPoint, endPoint);
            Acad.AddArcByStartAndAngle(centerPoint, startPoint, 90);
            Acad.AddArcByStartAndChordLength(centerPoint, startPoint, 70);
            Acad.AddArcByStartAndArcLength(centerPoint, startPoint, 200);
            Acad.AddArcBy3Point(startPoint, endPoint,centerPoint);//注意三点的顺序为逆时针方向
            Acad.Zoom();
        }
        /// <summary>
        /// 第11-12节课，创建举行和多边形
        /// </summary>
        public void P11to12()
        {
            double[] point1 = { 10, 10, 0 };
            double[] point2 = { 30, 20, 0 };
            Acad.AddRectangleBy2Point(point1, point2);
            Acad.AddPolygonInside(point1, point2, 5);
            Acad.AddPolygonOutside(point1, point2, 5);
            Acad.Zoom();
        }
        /// <summary>
        /// 第13节课，创建椭圆
        /// </summary>
        public void P13()
        {

            Acad.Zoom();
        }

    }
}
