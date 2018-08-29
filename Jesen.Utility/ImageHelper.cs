using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Jesen.Utility
{
    public class ImageHelper
    {
        private static string ImagePath = ConfigurationManager.AppSettings["ImagePath"];
        private static string VerifyPath = ConfigurationManager.AppSettings["ImagePath"];
        //绘图的原理很简单：Bitmap就像一张画布，Graphics如同画图的手，把Pen或Brush等绘图对象画在Bitmap这张画布上

        #region 验证码
        /// <summary>
        /// 画验证码
        /// </summary>
        public static void Drawing()
        {
            Bitmap bitmapobj = new Bitmap(100, 100);
            //在Bitmap上创建一个新的Graphics对象
            Graphics g = Graphics.FromImage(bitmapobj);
            //创建绘画对象，如Pen,Brush等
            Pen redPen = new Pen(Color.Red, 8);
            g.Clear(Color.White);
            //绘制图形
            g.DrawLine(redPen, 50, 20, 500, 20);
            g.DrawEllipse(Pens.Black, new Rectangle(0, 0, 200, 100));//画椭圆
            g.DrawArc(Pens.Black, new Rectangle(0, 0, 100, 100), 60, 180);//画弧线
            g.DrawLine(Pens.Black, 10, 10, 100, 100);//画直线
            g.DrawRectangle(Pens.Black, new Rectangle(0, 0, 100, 200));//画矩形
            g.DrawString("我爱北京天安门", new Font("微软雅黑", 12), new SolidBrush(Color.Red), new PointF(10, 10));//画字符串
            //g.DrawImage(

            if (!Directory.Exists(ImagePath))
                Directory.CreateDirectory(ImagePath);

            bitmapobj.Save(ImagePath + "pic1.jpg", ImageFormat.Jpeg);
            //释放所有对象
            bitmapobj.Dispose();
            g.Dispose();
        }

        /// <summary>
        /// 验证码
        /// </summary>
        public static void VerificationCode()
        {
            Bitmap bitmapobj = new Bitmap(300, 300);
            //在Bitmap上创建一个新的Graphics对象
            Graphics g = Graphics.FromImage(bitmapobj);
            g.DrawRectangle(Pens.Black, new Rectangle(0, 0, 150, 50));//画矩形
            g.FillRectangle(Brushes.White, new Rectangle(1, 1, 149, 49));
            g.DrawArc(Pens.Blue, new Rectangle(10, 10, 140, 10), 150, 90);//干扰线
            string[] arrStr = new string[] { "我", "们", "孝", "行", "白", "到", "国", "中", "来", "真" };
            Random r = new Random();
            int i;
            for (int j = 0; j < 4; j++)
            {
                i = r.Next(10);
                g.DrawString(arrStr[i], new Font("微软雅黑", 15), Brushes.Red, new PointF(j * 30, 10));
            }
            bitmapobj.Save(Path.Combine(VerifyPath, "Verif.jpg"), ImageFormat.Jpeg);
            bitmapobj.Dispose();
            g.Dispose();
        }

        /// <summary>
        /// 获取验证码
        /// </summary>
        /// <returns></returns>
        public static MemoryStream GetVerificationCode()
        {
            Bitmap bitmapobj = new Bitmap(300, 300);
            //在Bitmap上创建一个新的Graphics对象
            Graphics g = Graphics.FromImage(bitmapobj);
            g.DrawRectangle(Pens.Black, new Rectangle(0, 0, 150, 50));//画矩形
            g.FillRectangle(Brushes.White, new Rectangle(1, 1, 149, 49));
            g.DrawArc(Pens.Blue, new Rectangle(10, 10, 140, 10), 150, 90);//干扰线
            string[] arrStr = new string[] { "我", "们", "孝", "行", "白", "到", "国", "中", "来", "真" };
            Random r = new Random();
            int i;
            for (int j = 0; j < 4; j++)
            {
                i = r.Next(10);
                g.DrawString(arrStr[i], new Font("微软雅黑", 15), Brushes.Red, new PointF(j * 30, 10));
            }
            bitmapobj.Save(Path.Combine(VerifyPath, "Verif.jpg"), ImageFormat.Jpeg);
            bitmapobj.Dispose();
            g.Dispose();

            MemoryStream ms = new MemoryStream();
            bitmapobj.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            g.Dispose();
            return ms;
        }
        #endregion

        #region 图片缩放
        /// <summary>
        /// 按比例缩放,图片不会变形，会优先满足原图和最大长宽比例最高的一项
        /// </summary>
        /// <param name="oldPath"></param>
        /// <param name="newPath"></param>
        /// <param name="maxWidth"></param>
        /// <param name="maxHeight"></param>
        public static void CompressPercent(string oldPath, string newPath, int maxWidth, int maxHeight)
        {
            Image _sourceImg = Image.FromFile(oldPath);
            double _newW = (double)maxWidth;
            double _newH = (double)maxHeight;
            double percentWidth = (double)_sourceImg.Width > maxWidth ? (double)maxWidth : (double)_sourceImg.Width;

            if ((double)_sourceImg.Height * (double)percentWidth / (double)_sourceImg.Width > (double)maxHeight)
            {
                _newH = (double)maxHeight;
                _newW = (double)maxHeight / (double)_sourceImg.Height * (double)_sourceImg.Width;
            }
            else
            {
                _newW = percentWidth;
                _newH = (percentWidth / (double)_sourceImg.Width) * (double)_sourceImg.Height;
            }
            System.Drawing.Image bitmap = new System.Drawing.Bitmap((int)_newW, (int)_newH);
            Graphics g = Graphics.FromImage(bitmap);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.Clear(Color.Transparent);
            g.DrawImage(_sourceImg, new Rectangle(0, 0, (int)_newW, (int)_newH), new Rectangle(0, 0, _sourceImg.Width, _sourceImg.Height), GraphicsUnit.Pixel);
            _sourceImg.Dispose();
            g.Dispose();
            bitmap.Save(newPath, System.Drawing.Imaging.ImageFormat.Jpeg);
            bitmap.Dispose();
        }

        /// <summary>
        /// 按照指定大小对图片进行缩放，可能会图片变形
        /// </summary>
        /// <param name="oldPath"></param>
        /// <param name="newPath"></param>
        /// <param name="newWidth"></param>
        /// <param name="newHeight"></param>
        public static void ImageChangeBySize(string oldPath, string newPath, int newWidth, int newHeight)
        {
            Image sourceImg = Image.FromFile(oldPath);
            System.Drawing.Image bitmap = new System.Drawing.Bitmap(newWidth, newHeight);
            Graphics g = Graphics.FromImage(bitmap);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.Clear(Color.Transparent);
            g.DrawImage(sourceImg, new Rectangle(0, 0, newWidth, newHeight), new Rectangle(0, 0, sourceImg.Width, sourceImg.Height), GraphicsUnit.Pixel);
            sourceImg.Dispose();
            g.Dispose();
            bitmap.Save(newPath, System.Drawing.Imaging.ImageFormat.Jpeg);
            bitmap.Dispose();
        }

        #endregion

        #region 添加水印
        //添加水印
        /// <summary>
        /// 图片水印
        /// </summary>
        /// <param name="imgPath">服务器图片相对路径</param>
        /// <param name="filename">保存文件名</param>
        /// <param name="watermarkFilename">水印文件相对路径</param>
        /// <param name="watermarkStatus">图片水印位置 0=不使用 1=左上 2=中上 3=右上 4=左中  9=右下</param>
        /// <param name="quality">附加水印图片质量,0-100</param>
        /// <param name="watermarkTransparency">水印的透明度 1--10 10为不透明</param>
        public static void AddImageSignPic(string imgPath, string filename, string watermarkFilename, int watermarkStatus, int quality, int watermarkTransparency)
        {
            if (!File.Exists(GetMapPath(imgPath)))
                return;
            byte[] _ImageBytes = File.ReadAllBytes(GetMapPath(imgPath));
            Image img = Image.FromStream(new System.IO.MemoryStream(_ImageBytes));
            filename = GetMapPath(filename);

            if (watermarkFilename.StartsWith("/") == false)
                watermarkFilename = "/" + watermarkFilename;
            watermarkFilename = GetMapPath(watermarkFilename);
            if (!File.Exists(watermarkFilename))
                return;
            Graphics g = Graphics.FromImage(img);
            //设置高质量插值法
            //g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
            //设置高质量,低速度呈现平滑程度
            //g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            Image watermark = new Bitmap(watermarkFilename);

            if (watermark.Height >= img.Height || watermark.Width >= img.Width)
                return;

            ImageAttributes imageAttributes = new ImageAttributes();
            ColorMap colorMap = new ColorMap();

            colorMap.OldColor = Color.FromArgb(255, 0, 255, 0);
            colorMap.NewColor = Color.FromArgb(0, 0, 0, 0);
            ColorMap[] remapTable = { colorMap };

            imageAttributes.SetRemapTable(remapTable, ColorAdjustType.Bitmap);

            float transparency = 0.5F;
            if (watermarkTransparency >= 1 && watermarkTransparency <= 10)
                transparency = (watermarkTransparency / 10.0F);


            float[][] colorMatrixElements = {
                                                new float[] {1.0f,  0.0f,  0.0f,  0.0f, 0.0f},
                                                new float[] {0.0f,  1.0f,  0.0f,  0.0f, 0.0f},
                                                new float[] {0.0f,  0.0f,  1.0f,  0.0f, 0.0f},
                                                new float[] {0.0f,  0.0f,  0.0f,  transparency, 0.0f},
                                                new float[] {0.0f,  0.0f,  0.0f,  0.0f, 1.0f}
                                            };

            ColorMatrix colorMatrix = new ColorMatrix(colorMatrixElements);

            imageAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

            int xpos = 0;
            int ypos = 0;

            switch (watermarkStatus)
            {
                case 1:
                    xpos = (int)(img.Width * (float).01);
                    ypos = (int)(img.Height * (float).01);
                    break;
                case 2:
                    xpos = (int)((img.Width * (float).50) - (watermark.Width / 2));
                    ypos = (int)(img.Height * (float).01);
                    break;
                case 3:
                    xpos = (int)((img.Width * (float).99) - (watermark.Width));
                    ypos = (int)(img.Height * (float).01);
                    break;
                case 4:
                    xpos = (int)(img.Width * (float).01);
                    ypos = (int)((img.Height * (float).50) - (watermark.Height / 2));
                    break;
                case 5:
                    xpos = (int)((img.Width * (float).50) - (watermark.Width / 2));
                    ypos = (int)((img.Height * (float).50) - (watermark.Height / 2));
                    break;
                case 6:
                    xpos = (int)((img.Width * (float).99) - (watermark.Width));
                    ypos = (int)((img.Height * (float).50) - (watermark.Height / 2));
                    break;
                case 7:
                    xpos = (int)(img.Width * (float).01);
                    ypos = (int)((img.Height * (float).99) - watermark.Height);
                    break;
                case 8:
                    xpos = (int)((img.Width * (float).50) - (watermark.Width / 2));
                    ypos = (int)((img.Height * (float).99) - watermark.Height);
                    break;
                case 9:
                    xpos = (int)((img.Width * (float).99) - (watermark.Width));
                    ypos = (int)((img.Height * (float).99) - watermark.Height);
                    break;
            }

            g.DrawImage(watermark, new Rectangle(xpos, ypos, watermark.Width, watermark.Height), 0, 0, watermark.Width, watermark.Height, GraphicsUnit.Pixel, imageAttributes);

            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            ImageCodecInfo ici = null;
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.MimeType.IndexOf("jpeg") > -1)
                    ici = codec;
            }
            EncoderParameters encoderParams = new EncoderParameters();
            long[] qualityParam = new long[1];
            if (quality < 0 || quality > 100)
                quality = 80;

            qualityParam[0] = quality;

            EncoderParameter encoderParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qualityParam);
            encoderParams.Param[0] = encoderParam;

            if (ici != null)
                img.Save(filename, ici, encoderParams);
            else
                img.Save(filename);

            g.Dispose();
            img.Dispose();
            watermark.Dispose();
            imageAttributes.Dispose();
        }

        /// <summary>
        /// 文字水印
        /// </summary>
        /// <param name="imgPath">服务器图片相对路径</param>
        /// <param name="filename">保存文件名</param>
        /// <param name="watermarkText">水印文字</param>
        /// <param name="watermarkStatus">图片水印位置 0=不使用 1=左上 2=中上 3=右上 4=左中  9=右下</param>
        /// <param name="quality">附加水印图片质量,0-100</param>
        /// <param name="fontsize">字体大小</param>
        /// <param name="fontname">字体</param>
        public static void AddImageSignText(string imgPath, string filename, string watermarkText, int watermarkStatus, int quality, int fontsize, string fontname = "微软雅黑")
        {
            byte[] _ImageBytes = File.ReadAllBytes(GetMapPath(imgPath));
            Image img = Image.FromStream(new System.IO.MemoryStream(_ImageBytes));
            filename = GetMapPath(filename);

            Graphics g = Graphics.FromImage(img);
            Font drawFont = new Font(fontname, fontsize, FontStyle.Regular, GraphicsUnit.Pixel);
            SizeF crSize;
            crSize = g.MeasureString(watermarkText, drawFont);

            float xpos = 0;
            float ypos = 0;

            switch (watermarkStatus)
            {
                case 1:
                    xpos = (float)img.Width * (float).01;
                    ypos = (float)img.Height * (float).01;
                    break;
                case 2:
                    xpos = ((float)img.Width * (float).50) - (crSize.Width / 2);
                    ypos = (float)img.Height * (float).01;
                    break;
                case 3:
                    xpos = ((float)img.Width * (float).99) - crSize.Width;
                    ypos = (float)img.Height * (float).01;
                    break;
                case 4:
                    xpos = (float)img.Width * (float).01;
                    ypos = ((float)img.Height * (float).50) - (crSize.Height / 2);
                    break;
                case 5:
                    xpos = ((float)img.Width * (float).50) - (crSize.Width / 2);
                    ypos = ((float)img.Height * (float).50) - (crSize.Height / 2);
                    break;
                case 6:
                    xpos = ((float)img.Width * (float).99) - crSize.Width;
                    ypos = ((float)img.Height * (float).50) - (crSize.Height / 2);
                    break;
                case 7:
                    xpos = (float)img.Width * (float).01;
                    ypos = ((float)img.Height * (float).99) - crSize.Height;
                    break;
                case 8:
                    xpos = ((float)img.Width * (float).50) - (crSize.Width / 2);
                    ypos = ((float)img.Height * (float).99) - crSize.Height;
                    break;
                case 9:
                    xpos = ((float)img.Width * (float).99) - crSize.Width;
                    ypos = ((float)img.Height * (float).99) - crSize.Height;
                    break;
            }

            g.DrawString(watermarkText, drawFont, new SolidBrush(Color.White), xpos + 1, ypos + 1);
            g.DrawString(watermarkText, drawFont, new SolidBrush(Color.Black), xpos, ypos);

            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            ImageCodecInfo ici = null;
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.MimeType.IndexOf("jpeg") > -1)
                    ici = codec;
            }
            EncoderParameters encoderParams = new EncoderParameters();
            long[] qualityParam = new long[1];
            if (quality < 0 || quality > 100)
                quality = 80;

            qualityParam[0] = quality;

            EncoderParameter encoderParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qualityParam);
            encoderParams.Param[0] = encoderParam;

            if (ici != null)
                img.Save(filename, ici, encoderParams);
            else
                img.Save(filename);

            g.Dispose();
            img.Dispose();
        }

        #region 获得当前绝对路径
        /// <summary>
        /// 获得当前绝对路径
        /// </summary>
        /// <param name="strPath">指定的路径</param>
        /// <returns>绝对路径</returns>
        public static string GetMapPath(string strPath)
        {
            if (strPath.ToLower().StartsWith("http://"))
            {
                return strPath;
            }
            if (HttpContext.Current != null)
            {
                return HttpContext.Current.Server.MapPath(strPath);
            }
            else //非web程序引用
            {
                strPath = strPath.Replace("/", "\\");
                if (strPath.StartsWith("\\"))
                {
                    strPath = strPath.Substring(strPath.IndexOf('\\', 1)).TrimStart('\\');
                }
                return System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, strPath);
            }
        }
        #endregion 

        #endregion

        //二维码 QRCode
    }
}
