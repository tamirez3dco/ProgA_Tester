using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

namespace WebApplication1
{
    public partial class TableForm : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string path = HttpContext.Current.Server.MapPath("~/App_Data");
            DirectoryInfo din = new DirectoryInfo(path);
            DirectoryInfo[] hw_dirs = din.GetDirectories();
            HtmlTableRow tr = new HtmlTableRow();
            foreach (DirectoryInfo hw_dir in hw_dirs)
            {
                Object id = Session["myxxx"];
                int num = int.Parse((String)id);
                FileInfo[] fins = hw_dir.GetFiles("*" + num.ToString() + "*");
                if (fins.Length > 0)
                {
                    HtmlTableCell tc = new HtmlTableCell();
                    HyperLink hl = new HyperLink();
                    //?file = abc.txt
                    hl.NavigateUrl = "filedownload.ashx?file="+fins[0].FullName;
                    hl.Text = "Download " + hw_dir.Name;
                    tc.Controls.Add(hl);
                    tr.Cells.Add(tc);
                    Table1.Rows.Add(tr);
                }
                if (tr.Cells.Count == 0) Response.Write("Could not locate any HW for id=" + num);
            }

        }
    }
}