using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JasonAlgWeb
{
    public partial class myipaddr : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string access_ip = Request["REMOTE_ADDR"];
            IPHostEntry aaa = Dns.GetHostEntry(Dns.GetHostName());
                Response.Write("<br/><br/><br/><br/><br/><br/><br/><br/><br/><br/>"); 
                Response.Write("Your IP Address : " + access_ip + "<br/>");
                Response.Write("Your IP Address2 : " + getipadress() + "<br/>");
                /*
                Response.Write("SERVER IP Address List :<br/>");
                foreach (IPAddress ip in aaa.AddressList)
                {
                    Response.Write("server IP Address : " + ip.ToString() + "<br/>");
                }
                Response.Write("<br/>");
                Response.Write("server LOCAL IP Address : " + Request["LOCAL_ADDR"] + "<br/>");
                Response.Write("<br/>");
                 * */
        }
        private string getipadress()
        {
            System.Web.HttpContext context = System.Web.HttpContext.Current;
            if (context.Request["HTTP_X_FORWARDED_FOR"] == null)
            {
                return Request["REMOTE_ADDR"];
            }
            else
            {
                string[] iparray = context.Request["HTTP_X_FORWARDED_FOR"].Split(new Char[] { ',' });
                return iparray[0];
            }
        }
    }
}