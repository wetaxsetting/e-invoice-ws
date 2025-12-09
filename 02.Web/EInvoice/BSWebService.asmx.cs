using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Services;
using System.Xml.Serialization;

namespace EInvoice
{
    /// <summary>
    /// Summary description for WebService1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]

    public class BSWebService : System.Web.Services.WebService
    {
        [WebMethod]
        public string IssueInvoice(string arg_XmlStr)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
            using (TextReader reader = new StringReader(arg_XmlStr))
            {
                Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                if (_Invoices == null)
                {// chuyen toi login page

                }

            }
            return "{result:'OK'}";
        }
    }

    internal class TextReader
    {
    }

    internal class Invoices
    {
    }
}
