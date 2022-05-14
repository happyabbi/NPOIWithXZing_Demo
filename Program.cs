using System;
using System.Drawing;
using System.IO;
using NPOI.SS.UserModel;

namespace NPOIwithZXing_Demo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var targetFile = $@"{Environment.CurrentDirectory}/../../../Generated Excels/template.xlsx";

            var qrCode = BarcodeHelper.GenerateQRCode("Abraham Chen", 50, 50);
            var code39 = BarcodeHelper.GenerateCode39("Abraham Chen", 50, 100);
            var pdf417 = BarcodeHelper.GeneratePDF417("Abraham Chen", 300, 60);
            var dataMatrix = BarcodeHelper.GenerateDataMatrix("Abraham Chen", 50, 50);

            var x = LoadWorkbook(targetFile);

            var ws = x.CreateSheet("BarCode");

            var style = new PictureStyle
            {
                IsNoFill = false
            };


            var x1 = new PictureInfo(1, 4, 1, 2, ImageToByte(qrCode), style);
            AddPicture(ws, x1);

            var x2 = new PictureInfo(10, 14, 1, 4, ImageToByte(code39), style);
            AddPicture(ws, x2);

            var x3 = new PictureInfo(16, 18, 1, 4, ImageToByte(pdf417), style);
            AddPicture(ws, x3);

            var x4 = new PictureInfo(21, 24, 1, 2, ImageToByte(dataMatrix), style);
            AddPicture(ws, x4);

            var file = new FileStream($@"{Environment.CurrentDirectory}/../../../Generated Excels/barcode.xlsx",
                FileMode.Create); //產生檔案
            x.Write(file);
            file.Close();


            File.WriteAllBytes($"{Environment.CurrentDirectory}/../../../Generated images/qrCode.bmp",
                ImageToByte(qrCode));
            File.WriteAllBytes($"{Environment.CurrentDirectory}/../../../Generated images/code39.bmp",
                ImageToByte(code39));
            File.WriteAllBytes($"{Environment.CurrentDirectory}/../../../Generated images/pdf417.bmp",
                ImageToByte(pdf417));
            File.WriteAllBytes($"{Environment.CurrentDirectory}/../../../Generated images/dataMatrix.bmp",
                ImageToByte(dataMatrix));

        }

        public static void AddPicture(ISheet sheet, PictureInfo picInfo)
        {
            var pictureIdx = sheet.Workbook.AddPicture(picInfo.PictureData, PictureType.PNG);
            var anchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = picInfo.Col1;
            anchor.Col2 = picInfo.Col2;
            anchor.Row1 = picInfo.Row1;
            anchor.Row2 = picInfo.Row2;
            anchor.Dx1 = picInfo.PicturesStyle.AnchorDx1;
            anchor.Dx2 = picInfo.PicturesStyle.AnchorDx2;
            anchor.Dy1 = picInfo.PicturesStyle.AnchorDy1;
            anchor.Dy2 = picInfo.PicturesStyle.AnchorDy2;
            anchor.AnchorType = AnchorType.MoveDontResize;
            var drawing = sheet.CreateDrawingPatriarch();
            var pic = drawing.CreatePicture(anchor, pictureIdx);
            pic.Resize();
        }

        public static byte[] ImageToByte(Image img)
        {
            var converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }

        public static IWorkbook LoadWorkbook(string filePath)
        {
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return WorkbookFactory.Create(fileStream);
        }
    }
}
