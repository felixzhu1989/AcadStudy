using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Collections.Generic;

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
            Acad.AddArcBy3Point(startPoint, endPoint, centerPoint);//注意三点的顺序为逆时针方向
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
            double[] point4 = { 30, 0, 0 };
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
            AcadText text = Acad.AddText(textString, insertPoint, height);
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
            double[] points = new double[10];
            Acad.AddStarInCircleDemo(out points);
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
            dic.Add("A", "10");
            dic.Add("B", "20");
            Console.WriteLine("Value added for key = \"B\": {0}", dic["B"]);
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
        /// 第32-40课，小插件：批量修改器
        /// </summary>
        public void P32to40()
        {
            /*思路：
             * 提取文字，判断格式，加入选择集
             * 归类，计算修改后的参数
             * 循环，执行文字替换
             * 
             */
            Acad.Zoom();
        }

        #endregion 图元属性

        #region 图元修改
        /// <summary>
        /// 第41课，移动，旋转，删除实体
        /// </summary>
        public void P41()
        {
            Acad.MoveEntityDemo();
            Acad.RotateEntityDemo();
            Acad.DeleteEntityDemo();
            Acad.Zoom();
        }

        /// <summary>
        /// 第42课，调用CAD命令
        /// </summary>
        public void P42()
        {
            //Acad.SendCommandToCADDemo();
            Acad.SendCommandToCADAddLineByUser();
            Acad.Zoom();
        }

        /// <summary>
        /// 第43课，调用CAD命令，传递图元，打断命令
        /// </summary>
        public void P43()
        {
            Acad.SendCommandToCADBreak();
            Acad.Zoom();
        }
        /// <summary>
        /// 第44课，调用CAD命令，修剪实例
        /// </summary>
        public void P44_1()
        {
            Acad.TrimStar();
            Acad.Zoom();
        }
        /// <summary>
        /// 第44课，调用CAD命令，延伸
        /// </summary>
        public void P44_2()
        {
            Acad.SendCommandToCADExtend();
            Acad.Zoom();
        }

        /// <summary>
        /// 第45课，复制实体，镜像实体，缩放实体
        /// </summary>
        public void P45()
        {
            //Acad.CopyEntityDemo();
            //Acad.MirrorEntityDemo();
            Acad.ScaleEntityDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第46课，阵列
        /// </summary>
        public void P46()
        {
            //Acad.ArrayRectangularDemo();
            Acad.ArrayPolarDemo();
            Acad.Zoom();
        }

        /// <summary>
        /// 第47课，多段线操作,连接多段线,创建封闭区域
        /// </summary>
        public void P47()
        {
            //Acad.JoinLWPolyline();
            Acad.CreateBoundary();
            Acad.Zoom();
        }
        /// <summary>
        /// 第48课，用凸度的方式绘制多段线圆弧
        /// </summary>
        public void P48()
        {
            //Acad.GetBulgeInLWPolylineDemo();
            Acad.AddPolylineByBulgeDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第49课，将直线和圆转换成多段线
        /// </summary>
        public void P49()
        {
            AcadLine line = Acad.AddLineDemo();
            Acad.LineToLWPolyline(line);
            AcadCircle circle = Acad.AddCircleDemo();
            Acad.CircleToLWPolyline(circle);
            Acad.Zoom();
            Acad.ShowEntity();
        }
        /// <summary>
        /// 第50课，将圆弧转换成多段线
        /// </summary>
        public void P50()
        {
            AcadArc arc = Acad.AddArcDemo();
            Acad.ArcToLWPolyline(arc);

            Acad.Zoom();
            Acad.ShowEntity();
        }
        #endregion 图元修改


        #region 图层设置
        /// <summary>
        /// 第51课，创建和删除图层
        /// </summary>
        public void P51()
        {
            AcadLayer centerLayer = Acad.AddLayer("中心线");
            //Acad.DeleteLayer("中心线");
            //centerLayer.Delete();
            Acad.RenameLayer("中心线", "文字");
            Acad.Zoom();
        }
        /// <summary>
        /// 第52课，图层属性设置
        /// </summary>
        public void P52()
        {
            Acad.AddLayer("中心线");
            Acad.LayerSettingDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第53课，批量删除空白图层
        /// </summary>
        public void P53()
        {
            Acad.DeleteBlankLayer();
            Acad.Zoom();
        }
        #endregion 图层设置

        #region 图块操作
        /// <summary>
        /// 第54课，创建块
        /// </summary>
        public void P54()
        {
            Acad.CreateBlockDemo();
            Acad.Zoom();
        }
        /// <summary>
        /// 第55课，选择图元并创建块,插入块
        /// </summary>
        public void P55()
        {
            Acad.CreateBlockBySelect();
            //double[] insertPoint = { 0, 0, 0 };
            //Acad.InsertBlock("块实例", insertPoint);
            Acad.Zoom();
        }

        /// <summary>
        /// 第56课，删除块
        /// </summary>
        public void P56()
        {
            Acad.DeleteBlock("选择后创建块");
            Acad.Zoom();
        }

        /// <summary>
        /// 第57课，创建带属性的块，并插入
        /// </summary>
        public void P57()
        {
            Acad.CreateAttBlockDemo();
            //double[] insertPoint = { 0, 0, 0 };
            //AcadBlockReference block = Acad.InsertBlock("带属性的块", insertPoint);
            //设置属性在VBA中可行，但是移植到C#中后不可用
            //object[] blockAtt=(object[]) block.GetAttributes();
            // blockAtt[0] = "ANT2-2F";
            // blockAtt[1] = "12.2dBm";
            Acad.Zoom();
        }
        /// <summary>
        /// 第58课，批量修改块属性
        /// </summary>
        public void P58()
        {
            //Acad.ShowEntity();
            Acad.ChangeBlockAtt();
            Acad.Zoom();
        }



        #endregion 图块操作




    }
}
