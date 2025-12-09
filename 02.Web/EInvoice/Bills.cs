using System;
using System.Collections.Generic;
using System.Web;
using System.Xml.Serialization;

namespace EInvoice
{
    public class Bills
    {
        [XmlElement("Bn")]
        public List<Bn> Bn { get; set; }
    }
    public class Bn
    {
        [XmlElement("key")]
        public string key { get; set; }

        [XmlElement("Bill")]
        public Bill Bill { get; set; }
    }
    public class Bill
    {
        [XmlElement("Tos_Scan_Bill_Pk")]
        public string Tos_Scan_Bill_Pk { get; set; }

        [XmlElement("tos_customer_pk")]
        public string tos_customer_pk { get; set; }

        [XmlElement("Customer_ID")]
        public string Customer_ID { get; set; }

        [XmlElement("Customer_Nm")]
        public string Customer_Nm { get; set; }

        [XmlElement("Tax_Code")]
        public string Tax_Code { get; set; }

        [XmlElement("Bill_No")]
        public string Bill_No { get; set; }

        [XmlElement("Store_Name")]
        public string Store_Name { get; set; }

        [XmlElement("Store_Code")]
        public string Store_Code { get; set; }

        [XmlElement("Branch_Qr_Code")]
        public string Branch_Qr_Code { get; set; }

        [XmlElement("Total_Payment")]
        public string Total_Payment { get; set; }

        [XmlElement("Payment_Method")]
        public string Payment_Method { get; set; }

        [XmlElement("Bill_Attach")]
        public string Bill_Attach { get; set; }

        [XmlElement("Email")]
        public string Email { get; set; }

        [XmlElement("Address")]
        public string Address { get; set; }

        [XmlElement("Pos_No")]
        public string Pos_No { get; set; }

        [XmlElement("Sale_Date")]
        public string Sale_Date { get; set; }

        [XmlElement("tei_company_pk")]
        public string tei_company_pk { get; set; }

        [XmlElement("Combine_YN")]
        public string Combine_YN { get; set; }

        [XmlElement("User_id")]
        public string User_id { get; set; }
    }

}