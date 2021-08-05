using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace AcadStudy
{
    //BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
    //------测试代码---------------------
    public class Chapter
    {
        private AcadHelper Acad = new AcadHelper();

        #region 创建图元
        
        /// <summary>
        /// 第1-4课,创建直线
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
        /// 第5课，创建多段线
        /// </summary>
        public void P5()
        {
            Acad.AddLWPolylineDemo();
            Acad.AddPolylineDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第6课，创建圆
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
        /// 第7课，通过三点创建圆
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
        /// 第8课，使用循环方法创建同心圆
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
        /// 第8课，使用循环方法创建阵列
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
        /// 第9-10课，创建圆弧
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
        /// 第11-12课，创建举行和多边形
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
        /// 第13课，创建椭圆
        /// </summary>
        public void P13()
        {
            double[] point1 = { 0, 0, 0 };
            double[] point2 = { 20, 0, 0 };
            double[] point3 = { 10, 0, 0 };
            double[] point4 = {30, 0, 0};
            Acad.AddEllipseDemo();
            Acad.AddEllipseByAxisAndRatio(point3, point2, 0.7);
            Acad.AddEllipseBy3Point(point1, point4, point3);
            Acad.Zoom();
        }
        /// <summary>
        /// 第14-15课,创建填充
        /// </summary>
        public void P14toP15()
        {
            Acad.AddHatchDemo();
            Acad.AddGradientHatchDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第16课,创建文字
        /// </summary>
        public void P16()
        {
            double[] insertPoint = { 10, 10, 0 };
            string textString = "AutoCAD C# 二次开发";
            double height = 10;
            AcadText text= Acad.AddText(textString, insertPoint, height);
            text.Rotation = 30 * Math.PI / 180;//旋转
            text.Height = 5;
            text.Backward = true;//水平翻转
            text.UpsideDown = true;//上下翻转
            text.Alignment = AcAlignment.acAlignmentCenter;//对齐方式
            double[] insertPoint2 = { 10, -10, 0 };
            double width = 60;
            AcadMText mText = Acad.AddMText(insertPoint2, width, textString);
            mText.Height = 6;//后期修改字高
            Acad.Zoom();
        }
        #endregion 创建图元 

        #region 图元属性

        /// <summary>
        /// 第17-18课,图元属性
        /// </summary>
        public void P17to18()
        {
            //Acad.ShowEntity();
            Acad.ChangeCircleDemo();
            Acad.ChangeTextDemo();
            Acad.Zoom();
        }

        /// <summary>
        /// 第19课,根据现有图元创建其他图元，给文字添加边框
        /// </summary>
        public void P19()
        {
            //Acad.ShowEntity();
            Acad.AddBoundingBoxOnText();
            Acad.Zoom();
        }
        /// <summary>
        /// 第20课，选择单个圆添加内接五角星
        /// </summary>
        public void P20()
        {
            Acad.AddStarInCircleDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第21课，选择集，全部图元
        /// </summary>
        public void P21()
        {
            Acad.SelectionSetAllEntity();
            Acad.Zoom();
        }
        /// <summary>
        /// 第22课，选择集，用户选择
        /// </summary>
        public void P22()
        {
            Acad.SelectionSetByUser();
            Acad.Zoom();
        }
        /// <summary>
        /// 第23课，选择集,过滤(未解决这个问题)
        /// </summary>
        public void P23()
        {
            Acad.SelectionSetByFilterCircle();
            Acad.Zoom();
        }
        /// <summary>
        /// 第24课，手动选择图元添加倒选择集（过滤单行文本），并改成红色
        /// </summary>
        public void P24()
        {
            //Acad.ShowEntity();
            Acad.ChangeEntityInSelSet();
            Acad.Zoom();
        }
        /// <summary>
        /// 第25课，对齐单行文本
        /// </summary>
        public void P25()
        {
            Acad.SortText();
            Acad.Zoom();
        }

        /// <summary>
        /// 第29课，字典类型
        /// </summary>
        public void P29()
        {
            //字典类型
            Dictionary<string, string> dic =
                new Dictionary<string, string>();
            dic.Add("A","10");
            dic.Add("B","20");
            Console.WriteLine("Value added for key = \"B\": {0}",dic["B"]);
            Acad.Zoom();
        }
        /// <summary>
        /// 第30课，字典类型应用，统计图元个数
        /// </summary>
        public void P30()
        {
            Acad.EntitiesStatistics();
            Acad.Zoom();
        }

        /// <summary>
        /// 第31课，小插件：材料统计器
        /// </summary>
        public void P31()
        {
            Acad.MaterialStatistics();
            Acad.Zoom();
        }
        /// <summary>
        /// 第32课，小插件：批量修改器
        /// </summary>
        public void P32to33()
        {
            
            Acad.Zoom();
        }


        #endregion 图元属性

    }
}
