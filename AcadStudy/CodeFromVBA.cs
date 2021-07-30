using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace AcadStudy
{
    /*《ActiveX 和 VBA 参考》由明经通道翻译并提供：https://blog.csdn.net/weixin_46656590/article/details/105247189 
    下载地址：https://yunpan.360.cn/surl_yRE3CKKEqLk （提取码：c16e）
    BiliBili视频：https://www.bilibili.com/video/BV1Hb411T7sA
    */

    public class CodeFromVBA
    {
        private AcadApplication acadApp = AcadConn.GetAcadApplication();
        /// <summary>
        /// 将当前视图缩放到图形界限。
        /// </summary>
        public void Zoom()
        {
            acadApp.ZoomExtents();
        }
        /// <summary>
        /// 绘制一条预定义的直线
        /// </summary>
        /// <returns></returns>
        public AcadLine AddLineByPreDefine()
        {
            double[] startPoint = { 0, 0, 0 };
            double[] endPoint = { 50, 50, 0 };
            AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(startPoint, endPoint);
            return line;
        }
        /// <summary>
        /// 根据起始和结束点绘制直线
        /// </summary>
        /// <param name="startPoint">起始点</param>
        /// <param name="endPoint">结束点</param>
        /// <returns></returns>
        public AcadLine AddLineByPoint(double[] startPoint, double[] endPoint)
        {
            AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(startPoint, endPoint);
            return line;
        }
        /// <summary>
        /// 根据起始和结束点坐标值绘制直线
        /// </summary>
        /// <param name="startX"></param>
        /// <param name="stratY"></param>
        /// <param name="endX"></param>
        /// <param name="endY"></param>
        /// <returns></returns>
        public AcadLine AddLineByXY(double startX, double stratY, double endX, double endY)
        {
            double[] startPoint = { startX, stratY, 0 };
            double[] endPoint = { endX, endY, 0 };
            AcadLine line = acadApp.ActiveDocument.ModelSpace.AddLine(startPoint, endPoint);
            return line;
        }








    }
}
