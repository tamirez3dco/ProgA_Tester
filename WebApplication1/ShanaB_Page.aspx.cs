using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication1
{
    public partial class ShanaB_Page : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Walla");
            Session["myxxx"] = TextBox1.Text;
            Session["MyClassName"] = "EDP";
            Response.Redirect("TableForm.aspx");
        }
    }
}