using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using ZXing;
using ZXing.Common;
using ZXing.Datamatrix;
using ZXing.Datamatrix.Encoder;
using ZXing.PDF417;
using ZXing.QrCode;
using ZXing.QrCode.Internal;

namespace NPOIwithZXing_Demo
{
    internal class BarcodeHelper
    {
        public static Bitmap GenerateQRCode(string text, int width, int height)
        {
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE
            };
            var options = new QrCodeEncodingOptions
            {
                DisableECI = true,
                CharacterSet = "UTF-8",
                Width = width,
                Height = height,
                Margin = 1
            };

            writer.Options = options;
            var map = writer.Write(text);
            return map;
        }

        public static Bitmap GeneratePDF417(string text, int width, int height)
        {
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.PDF_417
            };
            var options = new PDF417EncodingOptions()
            {
                Width = width,
                Height = height,
                PureBarcode = true
            };

            writer.Options = options;
            var map = writer.Write(text);
            return map;
        }


        public static Bitmap GenerateDataMatrix(string text, int width, int height)
        {
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.DATA_MATRIX
            };
            var options = new DatamatrixEncodingOptions()
            {
                //PureBarcode = true,
                Width = width,
                Height = height,
                SymbolShape = SymbolShapeHint.FORCE_SQUARE
            };

            writer.Options = options;
            var map = writer.Write(text);
            return map;
        }

        public static Bitmap GenerateCode39(string text, int width, int height)
        {
            var writer = new BarcodeWriter
            {
                Format = BarcodeFormat.CODE_39
            };
            var options = new EncodingOptions
            {
                Width = width,
                Height = height,
                Margin = 2
            };
            writer.Options = options;
            var map = writer.Write(text);
            return map;
        }



        public static Bitmap Generate3(string text, int width, int height)
        {
            //Logo 圖片
            var logoPath = AppDomain.CurrentDomain.BaseDirectory + @"\img\logo.png";
            var logo = new Bitmap(logoPath);
            //構造二維碼寫碼器
            var writer = new MultiFormatWriter();
            var hint = new Dictionary<EncodeHintType, object>
            {
                { EncodeHintType.CHARACTER_SET, "UTF-8" },
                { EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H }
            };

            //生成二維碼 
            var bm = writer.encode(text, BarcodeFormat.QR_CODE, width + 30, height + 30, hint);
            bm = DeleteWhite(bm);
            var barcodeWriter = new BarcodeWriter();
            var map = barcodeWriter.Write(bm);

            //獲取二維碼實際尺寸（去掉二維碼兩邊空白後的實際尺寸）
            var rectangle = bm.getEnclosingRectangle();

            //計算插入圖片的大小和位置
            var middleW = Math.Min(rectangle[2] / 3, logo.Width);
            var middleH = Math.Min(rectangle[3] / 3, logo.Height);
            var middleL = (map.Width - middleW) / 2;
            var middleT = (map.Height - middleH) / 2;

            var bmpimg = new Bitmap(map.Width, map.Height, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bmpimg);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.DrawImage(map, 0, 0, width, height);
            //白底將二維碼插入圖片
            g.FillRectangle(Brushes.White, middleL, middleT, middleW, middleH);
            g.DrawImage(logo, middleL, middleT, middleW, middleH);

            return bmpimg;
        }

        private static BitMatrix DeleteWhite(BitMatrix matrix)
        {
            var rec = matrix.getEnclosingRectangle();
            var resWidth = rec[2] + 1;
            var resHeight = rec[3] + 1;

            var resMatrix = new BitMatrix(resWidth, resHeight);
            resMatrix.clear();
            for (var i = 0; i < resWidth; i++)
                for (var j = 0; j < resHeight; j++)
                    if (matrix[i + rec[0], j + rec[1]])
                        resMatrix[i, j] = true;
            return resMatrix;
        }
    }
}