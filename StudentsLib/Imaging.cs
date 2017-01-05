using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{
    public class Imaging
    {
        public static Bitmap CropImage(Bitmap source, Rectangle section)
        {
            // An empty bitmap which will hold the cropped image
            Bitmap bmp = new Bitmap(section.Width, section.Height);

            Graphics g = Graphics.FromImage(bmp);

            // Draw the given area (section) of the source image
            // at location 0,0 on the empty bitmap (bmp)
            g.DrawImage(source, 0, 0, section, GraphicsUnit.Pixel);

            return bmp;
        }
        public static double getSimilarity(Bitmap i1, Bitmap i2)
        {
            Bitmap smaller,larger;
            if (i1.Width > i2.Width)
            {
                larger = i1; smaller = i2;
            }
            else
            {
                larger = i2; smaller = i1;
            }
            Bitmap resized = new Bitmap(larger, smaller.Width, smaller.Height);
            double sum = 0;
            for (int c = 0; c < resized.Width; c++)
            {
                for (int r = 0; r < resized.Height; r++)
                {
                    Color p1 = resized.GetPixel(c,r);
                    Color p2 = smaller.GetPixel(c,r);
                    sum += Math.Abs(p1.R - p2.R);
                    sum += Math.Abs(p1.G - p2.G);
                    sum += Math.Abs(p1.B - p2.B);
                }
            }
            return sum / (resized.Width * resized.Height);
        }
    }
}
