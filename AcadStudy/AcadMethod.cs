using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace AcadStudy
{
    /// <summary>
    /// 参考：https://blog.csdn.net/shengshaohua/article/details/8363769
    /// </summary>
    public class AcadMethod
    {
        private AcadApplication acadApp = AcadConn.GetAcadApplication();
        public double Dimention_Line = 10;
        public AcadMethod() { }

        ///  /// <summary>
        /// 画竖立直线，带左边标注，从下至上
        /// </summary>
        /// <param name="StartX">起始点X值</param>
        /// <param name="StartY">起始点Y值</param>
        /// <param name="EndX">结束点X值</param>
        /// <param name="EndY">结束点Y值</param>
        /// <param name="WithDimention">是否带标注:true,false</param>
        public void DrawLineLeft(double StartX, double StartY, double EndX, double EndY, bool WithDimention)
        {
            if (WithDimention == false)
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
            }
            else
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
                if (StartX == EndX)
                {
                    AcadDimAligned aligned = acadApp.ActiveDocument.ModelSpace.AddDimAligned(new double[3] { StartX - Dimention_Line, StartY, 0 }, new double[3] { EndX - Dimention_Line, EndY, 0 }, new double[3] { (StartX + EndX) / 2 - Dimention_Line, (StartY + EndY) / 2, 0 });
                    AcadAcCmColor color = (AcadAcCmColor)acadApp.ActiveDocument.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");//AutoCAD2010 对应 18
                    color.SetRGB(225, 0, 0);
                    aligned.TrueColor = color;
                    /*pyautocad在给填充赋予颜色或者给绘制的图形填充颜色时候acad.doc.Application.GetInterfaceObject()是很好的选择，命令网上教程一般都是用的CAD2014或更低版本作为案例，使用的是acad.doc.Application.GetInterfaceObject(“AutoCAD.AcCmColor.19”)，但是参数(“AutoCAD.AcCmColor.19”)是同CAD版本相关的，使用其他版本时候请使用以下参数。
                    CAD2007：AutoCAD.AcCmColor.17
                    CAD2014：AutoCAD.AcCmColor.19
                    CAD2015：AutoCAD.AcCmColor.20
                    CAD2016/17：AutoCAD.AcCmColor.21
                    CAD2018：AutoCAD.AcCmColor.22
                    CAD2019：AutoCAD.AcCmColor.23
                    */
                }
            }
        }



        /// <summary>
        ///  画竖立直线，带右边标注，从下至上
        /// </summary>
        /// <param name="StartX">起始点X值</param>
        /// <param name="StartY">起始点Y值</param>
        /// <param name="EndX">结束点X值</param>
        /// <param name="EndY">结束点Y值</param>
        /// <param name="WithDimention">是否带标注:true,false</param>

        public void DrawLineRight(double StartX, double StartY, double EndX, double EndY, bool WithDimention)
        {

            if (WithDimention == false)
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
            }
            else
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
                if (StartX == EndX)
                {
                    AcadDimAligned aligned = acadApp.ActiveDocument.ModelSpace.AddDimAligned(new double[3] { StartX + Dimention_Line, StartY, 0 }, new double[3] { EndX + Dimention_Line, EndY, 0 }, new double[3] { (StartX + EndX) / 2 + Dimention_Line, (StartY + EndY) / 2, 0 });
                    AcadAcCmColor color = (AcadAcCmColor)acadApp.ActiveDocument.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");//AutoCAD2010 对应 18
                    color.SetRGB(225, 0, 0);
                    aligned.TrueColor = color;
                }
            }

        }


        /// <summary>
        ///  画横直线，带上方标注，从左到右
        /// </summary>
        /// <param name="StartX">起始点X值</param>
        /// <param name="StartY">起始点Y值</param>
        /// <param name="EndX">结束点X值</param>
        /// <param name="EndY">结束点Y值</param>
        /// <param name="WithDimention">是否带标注:true,false</param>

        public void DrawLineUp(double StartX, double StartY, double EndX, double EndY, bool WithDimention)
        {
            if (WithDimention == false)
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
            }
            else
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
                if (StartY == EndY)
                {
                    AcadDimAligned aligned = acadApp.ActiveDocument.ModelSpace.AddDimAligned(new double[3] { StartX, StartY + Dimention_Line, 0 }, new double[3] { EndX, EndY + Dimention_Line, 0 }, new double[3] { (StartX + EndX) / 2, StartY + Dimention_Line, 0 });
                    AcadAcCmColor color = (AcadAcCmColor)acadApp.ActiveDocument.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");//AutoCAD2010 对应 18
                    color.SetRGB(225, 0, 0);
                    aligned.TrueColor = color;
                }
            }
        }

        /// <summary>
        /// 画横直线，带下标注，从左至右
        /// </summary>
        /// <param name="StartX">起始点X值</param>
        /// <param name="StartY">起始点Y值</param>
        /// <param name="EndX">结束点X值</param>
        /// <param name="EndY">结束点Y值</param>
        /// <param name="WithDimention">是否带标注:true,false</param>
        public void DrawLineDown(double StartX, double StartY, double EndX, double EndY, bool WithDimention)
        {

            if (WithDimention == false)
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
            }
            else
            {
                AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(new double[3] { StartX, StartY, 0 }, new double[3] { EndX, EndY, 0 });
                if (StartY == EndY)
                {
                    AcadDimAligned aligned = acadApp.ActiveDocument.ModelSpace.AddDimAligned(new double[3] { StartX, StartY - Dimention_Line, 0 }, new double[3] { EndX, EndY - Dimention_Line, 0 }, new double[3] { (StartX + EndX) / 2, (StartY + EndY) / 2 - Dimention_Line, 0 });
                    AcadAcCmColor color = (AcadAcCmColor)acadApp.ActiveDocument.Application.GetInterfaceObject("AutoCAD.AcCmColor.17");//AutoCAD2010 对应 18
                    color.SetRGB(225, 0, 0);
                    aligned.TrueColor = color;
                }
            }

        }

        /// <summary>
        /// 画矩形，带标注
        /// </summary>
        /// <param name="StartX">起始点X值</param>
        /// <param name="StartY">起始点Y值</param>
        /// <param name="Width">宽度</param>
        /// <param name="Height">高度</param>
        /// <param name="WithDimention">是否带标注:true,false</param>
        public void DrawRectangle(double StartX, double StartY, double Width, double Height, bool WithDimention)
        {
            DrawLineLeft(StartX, StartY, StartX, StartY + Height, WithDimention);
            DrawLineRight(StartX + Width, StartY, StartX + Width, StartY + Height, WithDimention);
            DrawLineUp(StartX, StartY + Height, StartX + Width, StartY + Height, WithDimention);
            DrawLineDown(StartX, StartY, StartX + Width, StartY, WithDimention);
        }


        /// <summary>
        /// 添加块
        /// </summary>
        /// <param name="StartX">插入起始点X值</param>
        /// <param name="StartY">插入起始点Y值</param>
        /// <param name="FilePath">块文件路径</param>
        public void AddBlock(double StartX, double StartY, string FilePath)
        {
            AcadBlockReference block = acadApp.ActiveDocument.ModelSpace.InsertBlock(new double[3] { StartX, StartY, 0 }, FilePath, 1, 1, 1, 0);
        }





    }
}
