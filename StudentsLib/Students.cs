using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace StudentsLib
{

    public class Student
    {

        public String email;
        public int id;
        public String first_name;
        public String last_name;

        public void Send_Gmail(String subject, String Body, String[] attachments)
        {
            var client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("proga.netanya@gmail.com", "yardena12"),
                EnableSsl = true
            };
            var message = new MailMessage("proga.netanya@gmail.com", email, subject, Body);
            foreach (String file in attachments) message.Attachments.Add(new Attachment(file));
            client.Send(message);
        }
    }

    public class Students
    {
        public static String students_Excel_path = @"D:\Tamir\Netanya_ProgrammingA\2017\students_name_id.xlsx";
        public static Dictionary<int, Student> students_dic;
        public Students()
        {
            students_dic = Exceler.getStudents(@"D:\Tamir\Netanya_ProgrammingA\2017\students_name_id.xlsx");
        }
    }
}
