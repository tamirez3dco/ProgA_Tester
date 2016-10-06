﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{
    public class RunResults
    {
        public int grade = 100;
        public List<String> error_lines = new List<string>();

        public String errorsAsSingleString()
        {
            String res = String.Empty;
            for (int i = 0; i < error_lines.Count; i++) res += (error_lines[i] + "\n");
            return res;
        }

        // overload operator +
        public static RunResults operator +(RunResults a, RunResults b)
        {
            RunResults rr = new RunResults();
            rr.grade -= (100 - a.grade);
            rr.grade -= (100 - b.grade);
            rr.error_lines.AddRange(a.error_lines);
            rr.error_lines.AddRange(b.error_lines);

            return rr;
        }
    }
}
