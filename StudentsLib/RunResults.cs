using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{
    public enum Source
    {
        INPUT = 1,
        OUTPUT = 2,
        ERROR = 3
    }
    public class RunLine
    {
        public Source s;
        public String text;
        public RunLine(Source _s, String _t)
        {
            s = _s;
            text = _t;
        }
        public override string ToString()
        {
            return String.Format("{0}:{1}", s.ToString(), text);
        }

        public static string GetErrors(List<RunLine> lines)
        {
            String res = String.Empty;
            foreach (RunLine line in lines)
            {
                if (line.s == Source.ERROR) res += (line.text + "\n");
            }
            return res;
        }

        public static string GetOutputs(List<RunLine> lines)
        {
            String res = String.Empty;
            foreach (RunLine line in lines)
            {
                if (line.s == Source.OUTPUT) res += (line.text + "\n");
            }
            return res;
        }
    }

    public class RunResults
    {
        public int grade = 100;
        public List<String> error_lines = new List<string>();
        public List<String> filesToAttach = new List<string>();

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

            foreach (String file in b.filesToAttach)
            {
                if (!rr.filesToAttach.Contains(file)) rr.filesToAttach.Add(file);
            }
            foreach (String file in a.filesToAttach)
            {
                if (!rr.filesToAttach.Contains(file)) rr.filesToAttach.Add(file);
            }

            return rr;
        }
    }
}
