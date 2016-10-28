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

        public void Send_Gmail(String subject, String Body, List<String> attachments)
        {
            var client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("proga.netanya@gmail.com", "yardena12"),
                EnableSsl = true
            };
            var message = new MailMessage("proga.netanya@gmail.com", email, subject, Body);
            //var message = new MailMessage("proga.netanya@gmail.com", "tamirlevi123@gmail.com", subject, Body);
            foreach (String file in attachments)
            {
                if (file == null || file == String.Empty) continue;
                message.Attachments.Add(new Attachment(file));
            }
            client.Send(message);
        }
    }

    public class Students
    {
        public static Dictionary<int, Student> students_dic;
        public Students(String filePath)
        {
            students_dic = Exceler.getStudents(filePath);
        }
        public Students(bool b)
        {
            Student tl = new Student();
            tl.first_name = "תמיר";
            tl.last_name = "לוי";
            tl.id = 029046117;
            tl.email = "tamirlevi123@gmail.com";
            students_dic = new Dictionary<int, Student>();
            students_dic[tl.id] = tl;
        }
    }
}
