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
