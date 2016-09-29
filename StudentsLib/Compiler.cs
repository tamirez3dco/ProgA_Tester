using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Diagnostics;
using SharpCompress.Reader;
using SharpCompress.Common;

namespace StudentsLib
{
    public class Compiler
    {
        public static String errorReason;


        public static bool BuildZippedProject(String path, out String resulting_exe_file_path)
        {
            resulting_exe_file_path = null;
            FileInfo file = new FileInfo(path);
            String extractionPath = file.FullName.Substring(0, file.FullName.Length - 4) + "_extracted";
            // unzipping
            DirectoryInfo din = Directory.CreateDirectory(extractionPath);
            //ZipFile.ExtractToDirectory(path,extractionPath);

            using (Stream stream = File.OpenRead(path))
            {
                var reader = ReaderFactory.Open(stream);
                while (reader.MoveToNextEntry())
                {
                    if (!reader.Entry.IsDirectory)
                    {
                        //Console.WriteLine(reader.Entry.Key);
                        reader.WriteEntryToDirectory(extractionPath, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
                    }
                }
            }


            // search for .sln file
            FileInfo[] solutionFiles = din.GetFiles("*.sln",SearchOption.AllDirectories);
            if (solutionFiles.Length > 1) {
                errorReason = "More then 1 solution";
                return false;
            }
            if (solutionFiles.Length < 1)
            {
                errorReason = "No solutions found";
                return false;
            }

            //compiling...
            ProcessStartInfo psi = new ProcessStartInfo(@"c:\Program Files (x86)\MSBuild\14.0\Bin\MSBuild.exe", solutionFiles[0].FullName +" /t:rebuild");
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.WorkingDirectory = @"c:\Program Files (x86)\MSBuild\14.0\Bin";
            Process p = Process.Start(psi);

            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            String buildLogPath = din.FullName + @"\buildlog.txt";
            File.WriteAllText(buildLogPath, output);

            // 
            String buildFailed = "Build FAILED";
            if (output.Contains(buildFailed))
            {
                int lastLocation = output.LastIndexOf(buildFailed);
                errorReason = "BuildFailed" + "\n" + output.Substring(lastLocation);
                return false;
            }
            String[] allBuildLines = File.ReadAllLines(buildLogPath);
            bool buildPathFound = false;
            for (int lineNum = allBuildLines.Length -1; lineNum >= 0; lineNum--)
            {
                String line = allBuildLines[lineNum];
                String bingo = "->";
                if (!line.Contains(bingo)) continue;
                int location = line.LastIndexOf(bingo);
                String exePath = line.Substring(location + 2).Trim();
                if (!File.Exists(exePath)) continue;
                else
                {
                    buildPathFound = true;
                    resulting_exe_file_path = exePath;
                    break;
                }
            }

            if (!buildPathFound)
            {
                errorReason = "Build Maybe Succeeded but no build path found";
                return false;
            }

            return true;
        }


    }
}
