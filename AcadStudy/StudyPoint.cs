using Autodesk.AutoCAD.Interop.Common;

namespace AcadStudy
{
    //BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
    //------测试代码---------------------
    public class StudyPoint
    {
        private AcadHelper Acad = new AcadHelper();
        /// <summary>
        /// 第1-4节课,创建直线
        /// </summary>
        public void P1_4()
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
            double[] centerPoint = {0, 0, 0};
            double[] firstPoint = { 50, 0, 0 };
            double[] secondPoint = {150, 0, 0};
            Acad.AddCircleDemo();
            AcadCircle railCircle = Acad.AddCircleByRadius(centerPoint, 100);
           Acad.AddCircleByDiameter(centerPoint, 300);
            AcadCircle moveCircle= Acad.AddCircleBy2Point(firstPoint, secondPoint);
            railCircle.color = ACAD_COLOR.acRed;
            moveCircle.color = ACAD_COLOR.acGreen;

            Acad.Zoom();
        }

        /// <summary>
        /// 第7节课，通过三点创建圆
        /// </summary>
        public void P7()
        {

            Acad.Zoom();
        }

    }
}
