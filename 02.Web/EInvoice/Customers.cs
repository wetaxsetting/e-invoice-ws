using System;
using System.Collections.Generic;
using System.Web;
using System.Xml.Serialization;

namespace EInvoice
{
    [XmlRoot("Customers")]
    public class Customers
    {
        [XmlElement("Cus")]
        public List<Cus> Cus { get; set; }
    }
    public class Cus
    {
        [XmlElement("key")]
        public string key { get; set; }

        [XmlElement("Customer")]
        public Customer Customer { get; set; }
    }

    public class Customer
    {
        [XmlElement("Cus_cd")]
        public string Cus_cd { get; set; }

        [XmlElement("Cus_nm")]
        public string Cus_nm { get; set; }

        [XmlElement("Cus_lnm")]
        public string Cus_lnm { get; set; }

        [XmlElement("Cus_fnm")]
        public string Cus_fnm { get; set; }

        [XmlElement("Tax_code")]
        public string Tax_code { get; set; }

        [XmlElement("Address_vn")]
        public string Address_vn { get; set; }

        [XmlElement("Address_en")]
        public string Address_en { get; set; }

        [XmlElement("Address_kr")]
        public string Address_kr { get; set; }

        [XmlElement("Phone")]
        public string Phone { get; set; }

        [XmlElement("Fax")]
        public string Fax { get; set; }

        [XmlElement("Email")]
        public string Email { get; set; }

        [XmlElement("Acc_no")]
        public string Acc_no { get; set; }

        [XmlElement("Acc_ccy")]
        public string Acc_ccy { get; set; }

        [XmlElement("Acc_holder")]
        public string Acc_holder { get; set; }

        [XmlElement("Bank_name")]
        public string Bank_name { get; set; }

        [XmlElement("Tei_company_pk")]
        public string Tei_company_pk { get; set; }

        [XmlElement("Remarks")]
        public string Remarks { get; set; }

        [XmlElement("Crt_by")]
        public string Crt_by { get; set; }

        [XmlElement("Web_site")]
        public string Web_site { get; set; }

        [XmlElement("Erp_customer_pk")]
        public string Erp_customer_pk { get; set; }

        [XmlElement("Buyer_name")]
        public string Buyer_name { get; set; }

        [XmlElement("User_login")]
        public string User_login { get; set; }

        [XmlElement("Company_tax_code")]
        public string Company_tax_code { get; set; }
        
        [XmlElement("Tax_code_To_UserID")]
        public string Tax_code_To_UserID { get; set; }

        [XmlElement("tei_customer_pk")]
        public string tei_customer_pk { get; set; }
    }


}