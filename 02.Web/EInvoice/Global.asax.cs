using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Security;
using System.Web.SessionState;

namespace EInvoice
{
    public class Global : System.Web.HttpApplication
    {

        protected void Application_Start(object sender, EventArgs e)
        {

        }

        protected void Session_Start(object sender, EventArgs e)
        {

        }

        protected void Application_BeginRequest(object sender, EventArgs e)
        {
            string[] allowedOrigin = new string[] { "http://genuclouding.com/wseinvoice/", "http://192.168.1.232/wseinvoice/" ,"http://192.168.60.242"};
            var origin = HttpContext.Current.Request.Headers["Origin"];
            //if (origin != null)
           // {
             //   for (int i = 0; i < allowedOrigin.Length; i++)
             //   {
               //     if (allowedOrigin[i].Equals(origin))
               //     {
						HttpContext.Current.Response.AddHeader("Access-Control-Allow-Origin", "*");
						//HttpContext.Current.Response.AddHeader("Access-Control-Allow-Methods", "GET,POST");
						//HttpContext.Current.Response.End();
				//	}
              //  }
           // }
        }

        protected void Application_AuthenticateRequest(object sender, EventArgs e)
        {

        }

        protected void Application_Error(object sender, EventArgs e)
        {

        }

        protected void Session_End(object sender, EventArgs e)
        {

        }

        protected void Application_End(object sender, EventArgs e)
        {

        }
    }
}