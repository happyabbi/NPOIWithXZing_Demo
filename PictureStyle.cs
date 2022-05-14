
using NPOI.SS.UserModel;

namespace NPOIwithZXing_Demo
{
    public class PictureStyle
    {
        public int AnchorDx1 { get; set; }
        public int AnchorDx2 { get; set; }
        public int AnchorDy1 { get; set; }
        public int AnchorDy2 { get; set; }

        public int FillColor { get; set; }
        public bool IsNoFill { get; set; }
        public LineStyle LineStyle { get; set; }
        public int LineStyleColor { get; set; }
        public double LineWidth { get; set; }
    }
}