using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{
    public class Exceler
    {

        public static Dictionary<int,Student> getStudents(String excelPath)
        {
            Dictionary<int, Student> res = new Dictionary<int, Student>();

            //CREATING OBJECTS OF WORD AND DOCUMENT
            Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook exWbk = exApp.Workbooks.Open(excelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet exWks = (Microsoft.Office.Interop.Excel.Worksheet)exWbk.Sheets["ציונים"];

            for (int row = 2; row <= 200; row++)
            {
                try
                {
                    var studentFirstName = (string)(exWks.Cells[row, 1] as Microsoft.Office.Interop.Excel.Range).Value;
                    var studentLastName = (string)(exWks.Cells[row, 2] as Microsoft.Office.Interop.Excel.Range).Value;
                    var studentId_str = (string)(exWks.Cells[row, 3] as Microsoft.Office.Interop.Excel.Range).Value;
                    if (studentId_str == null) break;
                    var student_email_str = (string)(exWks.Cells[row, 6] as Microsoft.Office.Interop.Excel.Range).Value;
                    int studentId = int.Parse((String)studentId_str);

                    Student student = new Student();
                    student.id = studentId;
                    student.email = student_email_str;
                    student.first_name = studentFirstName.Trim();
                    student.last_name = studentLastName.Trim();

                    res[studentId] = student;
                }
                catch (Exception e) {
                    break;
                }
            }
            exWbk.Close();
            exApp.Quit();
            return res;
        }

    }
}
