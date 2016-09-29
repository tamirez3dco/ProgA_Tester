using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace WebApplication1
{
    /// <summary>
    /// Summary description for filedownload
    /// </summary>
    public class filedownload : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            //context.Response.ContentType = "text/plain";
            //context.Response.Write("Hello World");
            context.Response.Clear();
            context.Response.ContentType = "application/octet-stream";
            String s = context.Request.QueryString["file"];
            FileInfo fin = new FileInfo(s);
            context.Response.AddHeader("Content-Disposition", "attachment; filename="+fin.Name);
            context.Response.WriteFile(s);
            context.Response.End();
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}