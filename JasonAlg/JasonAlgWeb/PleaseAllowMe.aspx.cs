using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JasonAlgWeb
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (Request["req"]!=null && Request["req"].ToString() == "JasonAlg")
                //Response.Write("HELLOJASON!!!&&&");
                Response.Write("NOWAY!!!&&&");
            else
            {
                Response.Write("HAHAHAHA <br/>");
                Response.Write("HAHAHAHA <br/>");
                Response.Write("HAHAHAHA <br/>");
                Response.Write("HAHAHAHA <br/>");
                Response.Write("&&&&&&");
            }
        }
    }
}
