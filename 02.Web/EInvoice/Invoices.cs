using System;
using System.Collections.Generic;
using System.Web;
using System.Xml.Serialization;

namespace EInvoice
{
    [XmlRoot("Invoices")]
    public class Invoices
    {
        [XmlElement("Inv")]
        public List<Inv> Inv { get; set; }
    }
    public class Inv
    {
        [XmlElement("key")]
        public string key { get; set; }

        [XmlElement("Invoice")]
        public Invoice Invoice { get; set; }
    }
    public class Invoice
    {
        [XmlElement("tei_company_pk")]
        public string tei_company_pk { get; set; }

        [XmlElement("company_name")]
        public string company_name { get; set; }

        [XmlElement("company_tax_code")]
        public string company_tax_code { get; set; }

        [XmlElement("TrsDate")]
        public string TrsDate { get; set; }

        [XmlElement("tr_ccy")]
        public string tr_ccy { get; set; }

        [XmlElement("Book_ccy")]
        public string Book_ccy { get; set; }

        [XmlElement("Book_exrate")]
        public string Book_exrate { get; set; }

        [XmlElement("Trs_exrate")]
        public string Trs_exrate { get; set; }

        [XmlElement("Account_pk")]
        public string Account_pk { get; set; }

        [XmlElement("Account_code")]
        public string Account_code { get; set; }

        [XmlElement("Account_nm")]
        public string Account_nm { get; set; }

        [XmlElement("Serial_no")]
        public string Serial_no { get; set; }

        [XmlElement("Form_No")]
        public string Form_No { get; set; }

        [XmlElement("Enclose")]
        public string Enclose { get; set; }

        [XmlElement("DeclNo")]
        public string DeclNo { get; set; }

        [XmlElement("DueDate")]
        public string DueDate { get; set; }

        [XmlElement("CusCode")]
        public string CusCode { get; set; }

        [XmlElement("CusName")]
        public string CusName { get; set; }

        [XmlElement("CusCom")]
        public string CusCom { get; set; }

        [XmlElement("Buyer")]
        public string Buyer { get; set; }

        [XmlElement("CusAddress")]
        public string CusAddress { get; set; }

        [XmlElement("CusPhone")]
        public string CusPhone { get; set; }

        [XmlElement("CusTaxCode")]
        public string CusTaxCode { get; set; }

        [XmlElement("Description")]
        public string Description { get; set; }

        [XmlElement("Local_Description")]
        public string Local_Description { get; set; }

        [XmlElement("Receiver")]
        public string Receiver { get; set; }

        [XmlElement("PaymentMethod")]
        public string PaymentMethod { get; set; }

        [XmlElement("GrossValue")]
        public string GrossValue { get; set; }

        [XmlElement("VatAmount0")]
        public string VatAmount0 { get; set; }

        [XmlElement("GrossValue0")]
        public string GrossValue0 { get; set; }

        [XmlElement("VatAmount5")]
        public string VatAmount5 { get; set; }

        [XmlElement("GrossValue5")]
        public string GrossValue5 { get; set; }

        [XmlElement("VatAmount10")]
        public string VatAmount10 { get; set; }

        [XmlElement("VATRate")]
        public string VATRate { get; set; }

        [XmlElement("GrossValue10")]
        public string GrossValue10 { get; set; }

        [XmlElement("Total")]
        public string Total { get; set; }

        [XmlElement("VATAmount")]
        public string VATAmount { get; set; }

        [XmlElement("Amount")]
        public string Amount { get; set; }

        [XmlElement("AmountInWords")]
        public string AmountInWords { get; set; }

        [XmlElement("ArisingDate")]
        public string ArisingDate { get; set; }

        [XmlElement("EmailDeliver")]
        public string EmailDeliver { get; set; }

        [XmlElement("SMSDeliver")]
        public string SMSDeliver { get; set; }

        [XmlElement("User_id")]
        public string User_id { get; set; }

        [XmlElement("tei_customer_pk")]
        public string tei_customer_pk { get; set; }

        [XmlElement("tac_abacctcode_pk_vat")]
        public string tac_abacctcode_pk_vat { get; set; }

        [XmlElement("Invoicedate")]
        public string Invoicedate { get; set; }

        [XmlElement("interface_itemcode_yn")]
        public string interface_itemcode_yn { get; set; }

        [XmlElement("tac_crca_pk")]
        public string tac_crca_pk { get; set; }

        [XmlElement("tr_type")]
        public string tr_type { get; set; }

        [XmlElement("Total_Trans")]
        public string Total_Trans { get; set; }

        [XmlElement("VATAmount_Trans")]
        public string VATAmount_Trans { get; set; }

        [XmlElement("Invoice_Type")]
        public string Invoice_Type { get; set; }

        [XmlElement("Invoice_Desc")]
        public string Invoice_Desc { get; set; }

        [XmlElement("Invoice_Desc2")]
        public string Invoice_Desc2 { get; set; }

        [XmlElement("Remark3")]
        public string Remark3 { get; set; }

        [XmlElement("DeclDate")]
        public string DeclDate { get; set; }

        [XmlElement("CusAddress1")]
        public string CusAddress1 { get; set; }

        [XmlElement("CusAddress2")]
        public string CusAddress2 { get; set; }

        [XmlElement("CusAddress3")]
        public string CusAddress3 { get; set; }

        [XmlElement("Products")]
        public Products Products { get; set; }

        [XmlElement("AfterAmount")]
        public string AfterAmount { get; set; }

        [XmlElement("BeforAmount")]
        public string BeforAmount { get; set; }

        [XmlElement("AdjustType")]
        public string AdjustType { get; set; }

        [XmlElement("GapAmount")]
        public string GapAmount { get; set; }
        

        [XmlElement("InvoiceRef")]
        public string InvoiceRef { get; set; }

        [XmlElement("InvoiceRef_Dt")]
        public string InvoiceRef_Dt { get; set; }

        [XmlElement("SerialNoRef")]
        public string SerialNoRef { get; set; }
        
        [XmlElement("FormNoRef")]
        public string FormNoRef { get; set; }

        [XmlElement("InvoicePkRef")]
        public string InvoicePkRef { get; set; }

        [XmlElement("Je_Source")]
        public string Je_Source { get; set; }

        [XmlElement("AdjustInfo")]
        public string AdjustInfo { get; set; }

        [XmlElement("ReplaceInfo")]
        public string ReplaceInfo { get; set; }

    }
    public class Products
    {
        [XmlElement("Product")]
        public List<Product> Product { get; set; }
    }

    public class Product
    {
        [XmlElement("Code")]
        public string Code { get; set; }

        [XmlElement("ProdName")]
        public string ProdName { get; set; }

        [XmlElement("ProdUnit")]
        public string ProdUnit { get; set; }

        [XmlElement("ProdQuantity")]
        public string ProdQuantity { get; set; }

        [XmlElement("ProdPrice")]
        public string ProdPrice { get; set; }

        [XmlElement("Total")]
        public string Total { get; set; }

        [XmlElement("Extra1")]
        public string Extra1 { get; set; }

        [XmlElement("VATAmount")]
        public string VATAmount { get; set; }

        [XmlElement("Amount")]
        public string Amount { get; set; }

        [XmlElement("DDescription")]
        public string DDescription { get; set; }

        [XmlElement("DLDescription")]
        public string DLDescription { get; set; }

        [XmlElement("tac_crcad_pk")]
        public string tac_crcad_pk { get; set; }

        [XmlElement("itemcode")]
        public string itemcode { get; set; }

        [XmlElement("tco_item_pk")]
        public string tco_item_pk { get; set; }

        [XmlElement("OrderNo")]
        public string OrderNo { get; set; }

        [XmlElement("Attribute_01")]
        public string Attribute_01 { get; set; }

        [XmlElement("Attribute_02")]
        public string Attribute_02 { get; set; }

        [XmlElement("Attribute_03")]
        public string Attribute_03 { get; set; }

        [XmlElement("Attribute_04")]
        public string Attribute_04 { get; set; }

        [XmlElement("Attribute_05")]
        public string Attribute_05 { get; set; }

        [XmlElement("Seq")]
        public string Seq { get; set; }

        [XmlElement("Seq_Dis")]
        public string Seq_Dis { get; set; }

    }
}