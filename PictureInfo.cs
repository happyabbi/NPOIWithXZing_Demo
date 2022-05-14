
using System;

namespace NPOIwithZXing_Demo
{
    public class PictureInfo
    {
        public PictureInfo(int row1, int row2, int col1, int col2, Byte[] pictureData, PictureStyle pictureStyle)
        {
            Row1 = row1;
            Row2 = row2;
            Col1 = col1;
            Col2 = col2;
            PictureData = pictureData;
            PicturesStyle = pictureStyle;
        }

        public int Row1 { get; set; }
        public int Row2 { get; set; }
        public int Col1 { get; set; }
        public int Col2 { get; set; }
        public Byte[] PictureData { get; set; }
        public PictureStyle PicturesStyle { get; set; }
    }
}