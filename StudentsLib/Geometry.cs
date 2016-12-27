using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{
    public class Segment
    {
        public PointF P0;
        public PointF P1;
        public Segment(PointF p0, PointF p1)
        {
            P0 = p0;
            P1 = p1;
        }

        public static double getDistnace(PointF p1, PointF p2)
        {
            return Math.Sqrt(Math.Pow(p1.X-p2.X,2)+ Math.Pow(p1.Y - p2.Y,2));
        }

        public static float[] getVector(PointF to, PointF from)
        {
            float[] res = new float[2];
            res[0] = to.X - from.X;
            res[1] = to.Y - from.Y;
            return res;
        }

        // dist_Point_to_Segment(): get the distance of a point to a segment
        //     Input:  a Point P and a Segment S (in any dimension)
        //     Return: the shortest distance from P to S
        public static double dist_Point_to_Segment(PointF P, Segment S)
        {
            float[] v = Segment.getVector(S.P1, S.P0);
            float[] w = Segment.getVector(P, S.P0);

            double c1 = Segment.dot(w, v);
            if (c1 <= 0)
                return Segment.getDistnace(P, S.P0);

            double c2 = Segment.dot(v, v);
            if (c2 <= c1)
                return Segment.getDistnace(P, S.P1);

            double b = c1 / c2;
            PointF Pb = new PointF((float)(S.P0.X + b * v[0]), (float)(S.P0.Y + b * v[1]));
            return Segment.getDistnace(P, Pb);
        }

        public static float dot(float[] v1, float[] v2)
        {
            
            float res = 0;
            for (int i = 0; i < v1.Length; i++)
            {
                res += v1[i] * v2[i];
            }
            return res;
        }
    }

    class Geometry
    {
    }
}
