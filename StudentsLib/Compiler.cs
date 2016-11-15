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
using System.Windows.Forms;

namespace StudentsLib
{
    public class Compiler
    {
        public static String errorReason;


        public static bool BuildZippedProject(String path, out String resulting_exe_file_path)
        {
            resulting_exe_file_path = null;
            bool isRarArchive = SharpCompress.Archive.Rar.RarArchive.IsRarFile(path);
            bool isZipArchive = SharpCompress.Archive.Zip.ZipArchive.IsZipFile(path);
            if (!(isRarArchive || isZipArchive))
            {
                errorReason = "What was uploaded is neither a Zip archive nor a Rar archive. Maybe you did not upload the entire Solution directory ?";
                return false;
            }
            
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

            // clean the obj directory...
            DirectoryInfo[] objDins = din.GetDirectories("obj", SearchOption.AllDirectories);
            if (objDins.Length == 1)
            {
                Directory.Delete(objDins[0].FullName, true);
                Thread.Sleep(1000);
            }

            // clean the bin directory...
            DirectoryInfo[] binDins = din.GetDirectories("bin", SearchOption.AllDirectories);
            if (objDins.Length == 1)
            {
                Directory.Delete(binDins[0].FullName, true);
                Thread.Sleep(1000);
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
            ProcessStartInfo psi = new ProcessStartInfo(@"c:\Program Files (x86)\MSBuild\14.0\Bin\MSBuild.exe", "\"" + solutionFiles[0].FullName + "\"" + " /t:rebuild");
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

            // next remove all Console.ReadKey() instructions from source and recompile...
            FileInfo[] csFiles = din.GetFiles("*.cs", SearchOption.AllDirectories);
            foreach (FileInfo csFile in csFiles)
            {
                String source = File.ReadAllText(csFile.FullName);
                String changedCode = source.Replace("Console.ReadKey()", "//Console.ReadKey()");
                File.WriteAllText(csFile.FullName, changedCode);
            }

            p = Process.Start(psi);

            output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            File.WriteAllText(buildLogPath, output);

            // 
            if (output.Contains(buildFailed))
            {
                MessageBox.Show("Rebuild failed while orig Build succeeded!!");
                int lastLocation = output.LastIndexOf(buildFailed);
                errorReason = "BuildFailed" + "\n" + output.Substring(lastLocation);
                return false;
            }
            allBuildLines = File.ReadAllLines(buildLogPath);
            buildPathFound = false;
            for (int lineNum = allBuildLines.Length - 1; lineNum >= 0; lineNum--)
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
                MessageBox.Show("Rebuild succeeded but no buildPathFound. " + errorReason);
                return false;
            }

            return true;
        }

        public static bool BuildJavaZippedProject(String path, out String resulting_exe_file_path)
        {
            resulting_exe_file_path = null;
            bool isRarArchive = SharpCompress.Archive.Rar.RarArchive.IsRarFile(path);
            bool isZipArchive = SharpCompress.Archive.Zip.ZipArchive.IsZipFile(path);
            if (!(isRarArchive || isZipArchive))
            {
                errorReason = "What was uploaded is neither a Zip archive nor a Rar archive. Maybe you did not upload the entire Solution directory ?";
                return false;
            }

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

            // clean the obj directory...
            FileInfo[] classFiles = din.GetFiles("*.class", SearchOption.AllDirectories);
            foreach (FileInfo classFile in classFiles)
            {
                File.Delete(classFile.FullName);
            }


            // search for .java files
            FileInfo[] javaFiles = din.GetFiles("*.java", SearchOption.AllDirectories);
            if (javaFiles.Length < 1)
            {
                errorReason = "No java files found";
                return false;
            }
            if (javaFiles.Length > 1)
            {
                FileInfo correctProgram = null;
                foreach (FileInfo javaFile in javaFiles)
                {
                    if (javaFile.Name.ToLower() == "program.java")
                    {
                        correctProgram = javaFile;
                        break;
                    }
                }
                if (correctProgram == null)
                {
                    errorReason = "Several .java files found but none of them is Program.java. Can not proceed in building...";
                    return false;
                }
                javaFiles[0] = correctProgram;
            }


            //compiling...
            ProcessStartInfo psi = new ProcessStartInfo(@"C:\Program Files\Java\jdk1.8.0_65\bin\javac.exe", " " + javaFiles[0].Name);
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.WorkingDirectory = javaFiles[0].Directory.FullName;
            Process p = Process.Start(psi);

            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            String buildLogPath = din.FullName + @"\buildlog.txt";
            File.WriteAllText(buildLogPath, output);

            // 
            String error = "error";
            if (output.Contains(error))
            {
                errorReason = "BuildFailed" + "\n" + output;
                return false;
            }

            String expectedClassName = javaFiles[0].Name.Replace(javaFiles[0].Extension, ".class");
            classFiles = din.GetFiles(expectedClassName, SearchOption.AllDirectories);
            if (classFiles.Length == 0)
            {
                errorReason = "Build Maybe Succeeded but no " + expectedClassName + " found";
                return false;
            }

            resulting_exe_file_path = classFiles[0].FullName;
            return true;
        }

    }
}
