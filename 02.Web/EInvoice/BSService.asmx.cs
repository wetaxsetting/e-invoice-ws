using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.Services;
using System.Xml.Serialization;
using System.Data;
using System.Data.OracleClient;
using System.Text;
using System.Xml;
using System.Xml.Schema;

namespace EInvoice
{
    /// <summary>
    /// Summary description for BSService
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]

    public class BSService : System.Web.Services.WebService
    {
        private string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
        // string connString = String.Format( "Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
        private string dbName = "NOBLANDBD", dbUser = "genuwin", dbPwd = "genuwin2";//NOBLANDBD  EINVOICE_252

        [WebMethod]
        public string IssueInvoice(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string p_tei_einvoice_m_pk_out = "", p_invoice_no_rtn = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                    if (_Invoices != null)
                    {



                        //  temp = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        string Procedure = "sp_upd_60110280_issue";

                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                        command.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_company_pk.ToString();
                        command.Parameters.Add("p_company_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_name.ToString();
                        command.Parameters.Add("p_company_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_tax_code.ToString();
                        command.Parameters.Add("p_customer_cd", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        command.Parameters.Add("p_customer_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusName.ToString();
                        command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_customer_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Account_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk_vat", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_abacctcode_pk_vat.ToString();
                        command.Parameters.Add("p_tr_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.TrsDate.ToString();
                        command.Parameters.Add("p_tr_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_serial_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Serial_no.ToString();
                        command.Parameters.Add("p_invoice_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoicedate.ToString();
                        command.Parameters.Add("p_form_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Form_No.ToString();
                        command.Parameters.Add("p_invoice_no", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_tr_ccy", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_ccy.ToString();
                        command.Parameters.Add("p_tr_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Trs_exrate.ToString();
                        command.Parameters.Add("p_tr_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_type.ToString();
                        command.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Description.ToString();
                        command.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Local_Description.ToString();
                        command.Parameters.Add("p_remark3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Remark3.ToString();
                        command.Parameters.Add("p_tot_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total_Trans.ToString();
                        command.Parameters.Add("p_tot_net_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total.ToString();
                        command.Parameters.Add("p_tot_net_tax_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total.ToString();
                        command.Parameters.Add("p_tot_vat_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount_Trans.ToString();
                        command.Parameters.Add("p_tot_vat_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount.ToString();
                        command.Parameters.Add("p_vat_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATRate.ToString();
                        command.Parameters.Add("p_pay_method", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.PaymentMethod.ToString();
                        command.Parameters.Add("p_printed_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_cancel_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_sample_yn", OracleType.VarChar, 1000).Value = "N";
                        command.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk;
                        command.Parameters.Add("p_invoice_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Type.ToString();
                        command.Parameters.Add("p_ei_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_addr1", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress1;
                        command.Parameters.Add("p_addr2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress2;
                        command.Parameters.Add("p_addr3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress3;
                        command.Parameters.Add("p_bk_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Book_exrate;
                        command.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn;
                        command.Parameters.Add("p_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusTaxCode;
                        command.Parameters.Add("p_invoice_desc", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc;
                        command.Parameters.Add("p_invoice_desc2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc2;
                        command.Parameters.Add("p_DECLARE_NO", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclNo;
                        command.Parameters.Add("p_DECLARE_DATE", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclDate;
                      //  command.Parameters.Add("p_Xml", OracleType.Clob,999999999).Value = arg_XmlStr;
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id;


                        command.Parameters.Add("p_tei_einvoice_m_pk_out", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;
                        command.Parameters.Add("p_invoice_no_rtn", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;


                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();

                        p_tei_einvoice_m_pk_out = command.Parameters["p_tei_einvoice_m_pk_out"].Value.ToString();
                        p_invoice_no_rtn = command.Parameters["p_invoice_no_rtn"].Value.ToString();

                        for (int i = 0; i < _Invoices.Inv[0].Invoice.Products.Product.Count; i++)
                        {

                            Procedure = " sp_upd_60110280_issue_d_v2";  //sp_upd_60110280_issue_d     sp_upd_60110280_issue_d_v2  sp_upd_60110280_issue_d_v3 (In Hoa)

                            OracleCommand command_d = new OracleCommand(Procedure, connection);
                            command_d.CommandType = CommandType.StoredProcedure;
                            command_d.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                            command_d.Parameters.Add("p_tei_einvoice_d_pk", OracleType.VarChar, 1000).Value = 0;
                            command_d.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = p_tei_einvoice_m_pk_out;
                            command_d.Parameters.Add("p_tco_item_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tco_item_pk.ToString();
                            command_d.Parameters.Add("p_qty", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdQuantity.ToString();
                            command_d.Parameters.Add("p_u_price", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdPrice.ToString();
                            command_d.Parameters.Add("p_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Amount.ToString();
                            command_d.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DDescription.ToString();
                            command_d.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DLDescription.ToString();
                            command_d.Parameters.Add("p_tac_crcad_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk.ToString();
                            command_d.Parameters.Add("p_item_uom", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdUnit.ToString();
                            command_d.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id.ToString();
                            command_d.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tac_crcad_pk.ToString();
                            command_d.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn.ToString();
                            command_d.Parameters.Add("p_item_code", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Code.ToString();
                            command_d.Parameters.Add("p_item_name", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdName.ToString();
                            command_d.Parameters.Add("p_OrderNo", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].OrderNo.ToString();
                            command_d.Parameters.Add("p_Seq", OracleType.VarChar, 1000).Value = i + 1;
                            command_d.Parameters.Add("p_Seq_Dis", OracleType.VarChar, 1000).Value = i + 1;

                            command_d.Parameters.Add("p_Att_01", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_01.ToString();
                            command_d.Parameters.Add("p_Att_02", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_02.ToString();
                            command_d.Parameters.Add("p_Att_03", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_03.ToString();
                            command_d.Parameters.Add("p_Att_04", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_04.ToString();
                            command_d.Parameters.Add("p_Att_05", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_05.ToString();



                            OracleDataAdapter da_d = new OracleDataAdapter(command_d);
                            command_d.ExecuteNonQuery();
                        }
                        connection.Close();
                        connection.Dispose();

                    }

                }
                return "{\"result\":" + p_tei_einvoice_m_pk_out + "," + "\"Invoice_no\":\"" + p_invoice_no_rtn + "\",\"msg\":\"OK\"}";
            }
            catch (Exception ex)
            {

                return ex.Message + arg_XmlStr;
            }
        }

        [WebMethod]
        public string ListIssueInvoice(string arg_XmlStr)
        {

            //   ESysLib.WriteLogError(arg_XmlStr);
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string p_tei_einvoice_m_pk_out = "", p_invoice_no_rtn = "", list_invoice = "", list_tei_einvoice = "", list_tac_crac = "";
            int count = 0;
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                    if (_Invoices != null)
                    {

                        for (int i = 0; i < _Invoices.Inv.Count; i++)
                        {
                            
                            string Procedure = "sp_upd_60110280_issue_list";

                            connection = new OracleConnection(_conString);
                            connection.Open();
                            OracleCommand command = new OracleCommand(Procedure, connection);
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                            command.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = 0;
                            command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tei_company_pk.ToString();
                            command.Parameters.Add("p_company_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.company_name.ToString();
                            command.Parameters.Add("p_company_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.company_tax_code.ToString();
                            command.Parameters.Add("p_customer_cd", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.CusCode.ToString();
                            command.Parameters.Add("p_customer_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.CusName.ToString();
                            command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tei_customer_pk.ToString();
                            command.Parameters.Add("p_tac_abacctcode_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Account_pk.ToString();
                            command.Parameters.Add("p_tac_abacctcode_pk_vat", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tac_abacctcode_pk_vat.ToString();
                            command.Parameters.Add("p_tr_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.TrsDate.ToString();
                            command.Parameters.Add("p_tr_status", OracleType.VarChar, 1000).Value = "";
                            command.Parameters.Add("p_serial_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Serial_no.ToString();
                            command.Parameters.Add("p_invoice_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Invoicedate.ToString();
                            command.Parameters.Add("p_form_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Form_No.ToString();
                            command.Parameters.Add("p_invoice_no", OracleType.VarChar, 1000).Value = "";
                            command.Parameters.Add("p_tr_ccy", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tr_ccy.ToString();
                            command.Parameters.Add("p_tr_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Trs_exrate.ToString();
                            command.Parameters.Add("p_tr_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tr_type.ToString();
                            command.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Description.ToString();
                            command.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Local_Description.ToString();
                            command.Parameters.Add("p_remark3", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Remark3.ToString();
                            command.Parameters.Add("p_tot_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Total_Trans.ToString();
                            command.Parameters.Add("p_tot_net_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Total.ToString();
                            command.Parameters.Add("p_tot_net_tax_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Total.ToString();
                            command.Parameters.Add("p_tot_vat_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.VATAmount_Trans.ToString();
                            command.Parameters.Add("p_tot_vat_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.VATAmount.ToString();
                            command.Parameters.Add("p_vat_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.VATRate.ToString();
                            command.Parameters.Add("p_pay_method", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.PaymentMethod.ToString();
                            command.Parameters.Add("p_printed_yn", OracleType.VarChar, 1000).Value = "";
                            command.Parameters.Add("p_cancel_yn", OracleType.VarChar, 1000).Value = "";
                            command.Parameters.Add("p_sample_yn", OracleType.VarChar, 1000).Value = "N";
                            command.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tac_crca_pk;
                            command.Parameters.Add("p_invoice_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Invoice_Type.ToString();
                            command.Parameters.Add("p_ei_status", OracleType.VarChar, 1000).Value = "";
                            command.Parameters.Add("p_addr1", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.CusAddress1;
                            command.Parameters.Add("p_addr2", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.CusAddress2;
                            command.Parameters.Add("p_addr3", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.CusAddress3;
                            command.Parameters.Add("p_bk_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Book_exrate;
                            command.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.interface_itemcode_yn;
                            command.Parameters.Add("p_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.CusTaxCode;
                            command.Parameters.Add("p_invoice_desc", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Invoice_Desc;
                            command.Parameters.Add("p_invoice_desc2", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Invoice_Desc2;
                            command.Parameters.Add("p_DECLARE_NO", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.DeclNo;
                            command.Parameters.Add("p_DECLARE_DATE", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.DeclDate;
                       //     command.Parameters.Add("p_Xml", OracleType.Clob, 999999).Value = "testtt";
                            command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.User_id;


                            command.Parameters.Add("p_tei_einvoice_m_pk_out", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;
                            command.Parameters.Add("p_invoice_no_rtn", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;


                            OracleDataAdapter da = new OracleDataAdapter(command);
                            command.ExecuteNonQuery();

                            p_tei_einvoice_m_pk_out = command.Parameters["p_tei_einvoice_m_pk_out"].Value.ToString();
                            p_invoice_no_rtn = command.Parameters["p_invoice_no_rtn"].Value.ToString();

                            list_invoice = p_invoice_no_rtn + "-" + list_invoice;
                            list_tei_einvoice = p_tei_einvoice_m_pk_out + "-" + list_tei_einvoice;
                            list_tac_crac = _Invoices.Inv[i].Invoice.tac_crca_pk + "-" + list_tac_crac;

                            count++;
                            for (int j = 0; j < _Invoices.Inv[i].Invoice.Products.Product.Count; j++)
                            {

                                Procedure = " sp_upd_60110280_issue_d_v2";  //sp_upd_60110280_issue_d     sp_upd_60110280_issue_d_v2  sp_upd_60110280_issue_d_v3 (In Hoa)

                                OracleCommand command_d = new OracleCommand(Procedure, connection);
                                command_d.CommandType = CommandType.StoredProcedure;
                                command_d.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                                command_d.Parameters.Add("p_tei_einvoice_d_pk", OracleType.VarChar, 1000).Value = 0;
                                command_d.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = p_tei_einvoice_m_pk_out;
                                command_d.Parameters.Add("p_tco_item_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].tco_item_pk.ToString();
                                command_d.Parameters.Add("p_qty", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].ProdQuantity.ToString();
                                command_d.Parameters.Add("p_u_price", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].ProdPrice.ToString();
                                command_d.Parameters.Add("p_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Amount.ToString();
                                command_d.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].DDescription.ToString();
                                command_d.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].DLDescription.ToString();
                                command_d.Parameters.Add("p_tac_crcad_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tac_crca_pk.ToString();
                                command_d.Parameters.Add("p_item_uom", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].ProdUnit.ToString();
                                command_d.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.User_id.ToString();
                                command_d.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].tac_crcad_pk.ToString();
                                command_d.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.interface_itemcode_yn.ToString();
                                command_d.Parameters.Add("p_item_code", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Code.ToString();
                                command_d.Parameters.Add("p_item_name", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].ProdName.ToString();
                                command_d.Parameters.Add("p_OrderNo", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].OrderNo.ToString();
                                command_d.Parameters.Add("p_Seq", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Seq.ToString();
                                command_d.Parameters.Add("p_Seq_Dis", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Seq_Dis.ToString();

                                command_d.Parameters.Add("p_Att_01", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Attribute_01.ToString();
                                command_d.Parameters.Add("p_Att_02", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Attribute_02.ToString();
                                command_d.Parameters.Add("p_Att_03", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Attribute_03.ToString();
                                command_d.Parameters.Add("p_Att_04", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Attribute_04.ToString();
                                command_d.Parameters.Add("p_Att_05", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.Products.Product[j].Attribute_05.ToString();

                                OracleDataAdapter da_d = new OracleDataAdapter(command_d);
                                command_d.ExecuteNonQuery();
                            }
                            connection.Close();
                            connection.Dispose();
                        }
                    }

                }
                return "{\"result\":" + '"' + list_tei_einvoice + '"' + "," + "\"TAC_CRAC_PK\":\"" + list_tac_crac + '"' + "," + "\"Count_EInvoice\":\"" + count + '"' + "," + "\"Invoice_no\":\"" + list_invoice + "\",\"msg\":\"OK\"}";
            }
            catch (Exception ex)
            {
                //ESysLib.WriteLogError(arg_XmlStr);
                string error = ex.Message.ToString();
                error = error.Replace(':', ' ').Replace('"', ' ').Replace(',', ' ').Replace('.', ' ').Replace('-', ' ').Replace('_', ' ').Substring(11);
                string[] array = error.Split('!');
                return "{\"result\":" + '"' + list_tei_einvoice + '"' + "," + "\"TAC_CRAC_PK\":\"" + list_tac_crac + '"' + "," + "\"Count_EInvoice\":\"" + count + '"' + "," + "\"Invoice_no\":\"" + list_invoice + "\",\"msg\":\"" + array[0].ToString() + "\"}";
                //  return "{\"result\":" + p_tei_einvoice_m_pk_out + "," + "\"Invoice_no\":\"" + p_invoice_no_rtn + "\",\"msg\":\"Error !!\"}";.0
                //return ex.Message;
            }
        }

        [WebMethod]
        public string IssueInvoice_v2(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string p_tei_einvoice_m_pk_out = "", p_invoice_no_rtn = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
                XmlTextReader readered = new XmlTextReader("books.xml");
                using (TextReader reader = new StringReader(arg_XmlStr))
                    
                {
                    Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                    if (_Invoices != null)
                    {

                        //  temp = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        string Procedure = "sp_upd_60110280_issue";

                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                        command.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_company_pk.ToString();
                        command.Parameters.Add("p_company_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_name.ToString();
                        command.Parameters.Add("p_company_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_tax_code.ToString();
                        command.Parameters.Add("p_customer_cd", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        command.Parameters.Add("p_customer_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusName.ToString();
                        command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_customer_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Account_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk_vat", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_abacctcode_pk_vat.ToString();
                        command.Parameters.Add("p_tr_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.TrsDate.ToString();
                        command.Parameters.Add("p_tr_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_serial_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Serial_no.ToString();
                        command.Parameters.Add("p_invoice_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoicedate.ToString();
                        command.Parameters.Add("p_form_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Form_No.ToString();
                        command.Parameters.Add("p_invoice_no", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_tr_ccy", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_ccy.ToString();
                        command.Parameters.Add("p_tr_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Trs_exrate.ToString();
                        command.Parameters.Add("p_tr_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_type.ToString();
                        command.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Description.ToString();
                        command.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Local_Description.ToString();
                        command.Parameters.Add("p_remark3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Remark3.ToString();
                        command.Parameters.Add("p_tot_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total_Trans.ToString();
                        command.Parameters.Add("p_tot_net_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total.ToString();
                        command.Parameters.Add("p_tot_net_tax_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total.ToString();
                        command.Parameters.Add("p_tot_vat_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount_Trans.ToString();
                        command.Parameters.Add("p_tot_vat_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount.ToString();
                        command.Parameters.Add("p_vat_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATRate.ToString();
                        command.Parameters.Add("p_pay_method", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.PaymentMethod.ToString();
                        command.Parameters.Add("p_printed_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_cancel_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_sample_yn", OracleType.VarChar, 1000).Value = "N";
                        command.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk;
                        command.Parameters.Add("p_invoice_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Type.ToString();
                        command.Parameters.Add("p_ei_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_addr1", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress1;
                        command.Parameters.Add("p_addr2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress2;
                        command.Parameters.Add("p_addr3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress3;
                        command.Parameters.Add("p_bk_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Book_exrate;
                        command.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn;
                        command.Parameters.Add("p_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusTaxCode;
                        command.Parameters.Add("p_invoice_desc", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc;
                        command.Parameters.Add("p_invoice_desc2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc2;
                        command.Parameters.Add("p_DECLARE_NO", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclNo;
                        command.Parameters.Add("p_DECLARE_DATE", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclDate;
                        command.Parameters.Add("p_Xml", OracleType.Clob, 999999).Value = arg_XmlStr;
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id;


                        command.Parameters.Add("p_tei_einvoice_m_pk_out", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;
                        command.Parameters.Add("p_invoice_no_rtn", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;


                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();

                        p_tei_einvoice_m_pk_out = command.Parameters["p_tei_einvoice_m_pk_out"].Value.ToString();
                        p_invoice_no_rtn = command.Parameters["p_invoice_no_rtn"].Value.ToString();

                        for (int i = 0; i < _Invoices.Inv[0].Invoice.Products.Product.Count; i++)
                        {

                            Procedure = " sp_upd_60110280_issue_d_v2";  //sp_upd_60110280_issue_d     sp_upd_60110280_issue_d_v2  sp_upd_60110280_issue_d_v3 (In Hoa)

                            OracleCommand command_d = new OracleCommand(Procedure, connection);
                            command_d.CommandType = CommandType.StoredProcedure;
                            command_d.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                            command_d.Parameters.Add("p_tei_einvoice_d_pk", OracleType.VarChar, 1000).Value = 0;
                            command_d.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = p_tei_einvoice_m_pk_out;
                            command_d.Parameters.Add("p_tco_item_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tco_item_pk.ToString();
                            command_d.Parameters.Add("p_qty", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdQuantity.ToString();
                            command_d.Parameters.Add("p_u_price", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdPrice.ToString();
                            command_d.Parameters.Add("p_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Amount.ToString();
                            command_d.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DDescription.ToString();
                            command_d.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DLDescription.ToString();
                            command_d.Parameters.Add("p_tac_crcad_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk.ToString();
                            command_d.Parameters.Add("p_item_uom", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdUnit.ToString();
                            command_d.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id.ToString();
                            command_d.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tac_crcad_pk.ToString();
                            command_d.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn.ToString();
                            command_d.Parameters.Add("p_item_code", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Code.ToString();
                            command_d.Parameters.Add("p_item_name", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdName.ToString();
                            command_d.Parameters.Add("p_OrderNo", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].OrderNo.ToString();
                            command_d.Parameters.Add("p_Seq", OracleType.VarChar, 1000).Value = i + 1;
                            command_d.Parameters.Add("p_Seq_Dis", OracleType.VarChar, 1000).Value = i + 1;

                            command_d.Parameters.Add("p_Att_01", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_01.ToString();
                            command_d.Parameters.Add("p_Att_02", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_02.ToString();
                            command_d.Parameters.Add("p_Att_03", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_03.ToString();
                            command_d.Parameters.Add("p_Att_04", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_04.ToString();
                            command_d.Parameters.Add("p_Att_05", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_05.ToString();



                            OracleDataAdapter da_d = new OracleDataAdapter(command_d);
                            command_d.ExecuteNonQuery();
                        }
                        //thêm vào{
                        while (readered.Read())
                        {
                            switch (readered.NodeType)
                            {
                                case XmlNodeType.Element: // The node is an element.
                                    Console.Write("<" + readered.Name);
                                    Console.WriteLine(">");
                                    break;
                                case XmlNodeType.Text: //Display the text in each element.
                                    Console.WriteLine(readered.Value);
                                    break;
                                case XmlNodeType.EndElement: //Display the end of the element.
                                    Console.Write("</" + readered.Name);
                                    Console.WriteLine(">");
                                    break;
                            }
                        }
                        ///}///
                        connection.Close();
                        connection.Dispose();

                    }

                }
                return "{\"result\":" + p_tei_einvoice_m_pk_out + "," + "\"Invoice_no\":\"" + p_invoice_no_rtn + "\",\"msg\":\"OK\"}";
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        [WebMethod]
        public string Cancel_Invoice(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string p_tei_einvoice_m_pk_out = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                    if (_Invoices != null)
                    {

                        //  temp = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        string Procedure = "ac_pro_60110320_cancel_einv";

                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_company_pk;
                        command.Parameters.Add("p_crda_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk;
                        command.Parameters.Add("p_index", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id.ToString();
                        command.Parameters.Add("p_rtn_value1", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;

                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();

                        p_tei_einvoice_m_pk_out = command.Parameters["p_rtn_value1"].Value.ToString();

                        connection.Close();
                        connection.Dispose();

                    }

                }
                return "{\"result\":" + p_tei_einvoice_m_pk_out + ",\"msg\":\"OK\"}";
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        [WebMethod]
        public string List_Cancel_Invoice(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string p_tei_einvoice_m_pk_out = "", list_tac_crac = "";
            int count = 0;
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                    if (_Invoices != null)
                    {

                        for (int i = 0; i < _Invoices.Inv.Count; i++)
                        {
                            //  temp = _Invoices.Inv[0].Invoice.CusCode.ToString();
                            string Procedure = "ac_pro_60110320_cancel_einv";

                            connection = new OracleConnection(_conString);
                            connection.Open();
                            OracleCommand command = new OracleCommand(Procedure, connection);
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.Add("p_crda_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.tac_crca_pk;
                            command.Parameters.Add("p_index", OracleType.VarChar, 1000).Value = 0;
                            command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[i].Invoice.User_id.ToString();
                            command.Parameters.Add("p_rtn_value1", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;

                            OracleDataAdapter da = new OracleDataAdapter(command);
                            command.ExecuteNonQuery();

                            p_tei_einvoice_m_pk_out = command.Parameters["p_rtn_value1"].Value.ToString();
                            list_tac_crac = _Invoices.Inv[i].Invoice.tac_crca_pk + "-" + list_tac_crac;
                            count++;
                            connection.Close();
                            connection.Dispose();
                        }
                    }

                }

                return "{\"result\":" + '"' + list_tac_crac + '"' + "," + "\"Count_Invoice\":\"" + count + '"' + ",\"msg\":\"OK\"}";
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        [WebMethod]
        public string Get_Status_EInvoice(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            int result = 0; string xml_json = "";
            OracleConnection connection;
            try
            {
                string Procedure = "ac_pro_getinfo_invoice_erp";

                connection = new OracleConnection(_conString);
                connection.Open();
                OracleCommand command = new OracleCommand(Procedure, connection);
                command.CommandType = CommandType.StoredProcedure;


                command.Parameters.Add("p_tco_company_pk", OracleType.VarChar, 1000).Value = arg_XmlStr;
                command.Parameters.Add("p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
                DataSet ds = new DataSet();
                OracleDataAdapter da = new OracleDataAdapter(command);
                da.Fill(ds);
                DataTable dt = ds.Tables[0];
                result = dt.Rows.Count;

                connection.Close();
                connection.Dispose();

                var JSONString = new StringBuilder();
                //if (dt.Rows.Count > 0)
                //{
                JSONString.Append("[");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    JSONString.Append("{");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j < dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\",");
                        }
                        else if (j == dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\"");
                        }
                    }
                    if (i == dt.Rows.Count - 1)
                    {
                        JSONString.Append("}");
                    }
                    else
                    {
                        JSONString.Append("},");
                    }
                }
                JSONString.Append("]");
                //}
                xml_json = JSONString.ToString();
                return xml_json;
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        [WebMethod]
        //[ScriptMethod(ResponseFormat = ResponseFormat.Json)]//Specify return format.
        public string InsertCustomers(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string result = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {

                XmlSerializer serializer = new XmlSerializer(typeof(Customers));

                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Customers _Customers = (Customers)serializer.Deserialize(reader);
                    if (_Customers == null)
                    {// chuyen toi login page

                    }
                    else
                    {

                        string Procedure = "sp_upd_agic00050_3";
                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                        command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_cus_cd", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_cd.ToString();
                        command.Parameters.Add("p_cus_nm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_nm.ToString();
                        command.Parameters.Add("p_cus_lnm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_lnm.ToString();
                        command.Parameters.Add("p_cus_fnm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_fnm.ToString();
                        command.Parameters.Add("p_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tax_code.ToString();
                        command.Parameters.Add("p_addr", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_en.ToString();
                        command.Parameters.Add("p_addr_l", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_vn.ToString();
                        command.Parameters.Add("p_addr_f", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_kr.ToString();
                        command.Parameters.Add("p_tel", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Phone.ToString();
                        command.Parameters.Add("p_fax", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Fax.ToString();
                        command.Parameters.Add("p_email", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Email.ToString();
                        command.Parameters.Add("p_acc_no", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_no.ToString();
                        command.Parameters.Add("p_acc_ccy", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_ccy.ToString();
                        command.Parameters.Add("p_acc_holder", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_holder.ToString();
                        command.Parameters.Add("p_bank_name", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Bank_name.ToString();
                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tei_company_pk.ToString();
                        command.Parameters.Add("p_remarks", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Remarks.ToString();
                        command.Parameters.Add("p_use_yn", OracleType.VarChar, 1000).Value = "Y";
                        command.Parameters.Add("p_web_site", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Web_site.ToString();
                        command.Parameters.Add("p_erp_customer_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Erp_customer_pk.ToString();
                        command.Parameters.Add("p_tco_buspartner_pk_itf", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Erp_customer_pk.ToString();
                        command.Parameters.Add("p_buyer_name", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Buyer_name.ToString();
                        command.Parameters.Add("p_user_login", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.User_login.ToString();
                        command.Parameters.Add("p_company_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Company_tax_code.ToString();
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Crt_by.ToString();
                        command.Parameters.Add("p_Tax_code_To_UserID", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tax_code_To_UserID.ToString();
                        command.Parameters.Add("p_rtn_master_pk", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;

                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();

                        result = command.Parameters["p_rtn_master_pk"].Value.ToString();
                        connection.Close();
                        connection.Dispose();
                        //return s;


                    }

                }

                return "{\"result\":" + result + ",\"msg\":\"OK\"}";



            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }

        [WebMethod]
        public string ListInsertCustomers(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string result = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {

                XmlSerializer serializer = new XmlSerializer(typeof(Customers));

                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Customers _Customers = (Customers)serializer.Deserialize(reader);
                    if (_Customers == null)
                    {// chuyen toi login page

                    }
                    else
                    {
                        for (int i = 0; i < _Customers.Cus.Count; i++)
                        {
                            string Procedure = "sp_upd_agic00050_3";
                            connection = new OracleConnection(_conString);
                            connection.Open();
                            OracleCommand command = new OracleCommand(Procedure, connection);
                            command.CommandType = CommandType.StoredProcedure;

                            command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                            command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = 0;
                            command.Parameters.Add("p_cus_cd", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Cus_cd.ToString();
                            command.Parameters.Add("p_cus_nm", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Cus_nm.ToString();
                            command.Parameters.Add("p_cus_lnm", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Cus_lnm.ToString();
                            command.Parameters.Add("p_cus_fnm", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Cus_fnm.ToString();
                            command.Parameters.Add("p_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Tax_code.ToString();
                            command.Parameters.Add("p_addr", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Address_en.ToString();
                            command.Parameters.Add("p_addr_l", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Address_vn.ToString();
                            command.Parameters.Add("p_addr_f", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Address_kr.ToString();
                            command.Parameters.Add("p_tel", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Phone.ToString();
                            command.Parameters.Add("p_fax", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Fax.ToString();
                            command.Parameters.Add("p_email", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Email.ToString();
                            command.Parameters.Add("p_acc_no", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Acc_no.ToString();
                            command.Parameters.Add("p_acc_ccy", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Acc_ccy.ToString();
                            command.Parameters.Add("p_acc_holder", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Acc_holder.ToString();
                            command.Parameters.Add("p_bank_name", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Bank_name.ToString();
                            command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Tei_company_pk.ToString();
                            command.Parameters.Add("p_remarks", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Remarks.ToString();
                            command.Parameters.Add("p_use_yn", OracleType.VarChar, 1000).Value = "Y";
                            command.Parameters.Add("p_web_site", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Web_site.ToString();
                            command.Parameters.Add("p_erp_customer_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Erp_customer_pk.ToString();
                            command.Parameters.Add("p_tco_buspartner_pk_itf", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Erp_customer_pk.ToString();
                            command.Parameters.Add("p_buyer_name", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Buyer_name.ToString();
                            command.Parameters.Add("p_user_login", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.User_login.ToString();
                            command.Parameters.Add("p_company_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Company_tax_code.ToString();
                            command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Crt_by.ToString();
                            command.Parameters.Add("p_Tax_code_To_UserID", OracleType.VarChar, 1000).Value = _Customers.Cus[i].Customer.Tax_code_To_UserID.ToString();
                            command.Parameters.Add("p_rtn_master_pk", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;

                            OracleDataAdapter da = new OracleDataAdapter(command);
                            command.ExecuteNonQuery();

                            result = command.Parameters["p_rtn_master_pk"].Value.ToString() + "-" + result;
                            connection.Close();
                            connection.Dispose();
                            //return s;
                        }
                    }

                }

                return "{\"result\":" +'"'+ result +'"'+ ",\"msg\":\"OK\"}";



            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }

        [WebMethod]
        public string UpdateCustomers(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string result = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {

                XmlSerializer serializer = new XmlSerializer(typeof(Customers));

                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Customers _Customers = (Customers)serializer.Deserialize(reader);
                    if (_Customers == null)
                    {// chuyen toi login page

                    }
                    else
                    {

                        string Procedure = "sp_upd_agic00050_3";
                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                        command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_cus_cd", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_cd.ToString();
                        command.Parameters.Add("p_cus_nm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_nm.ToString();
                        command.Parameters.Add("p_cus_lnm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_lnm.ToString();
                        command.Parameters.Add("p_cus_fnm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_fnm.ToString();
                        command.Parameters.Add("p_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tax_code.ToString();
                        command.Parameters.Add("p_addr", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_en.ToString();
                        command.Parameters.Add("p_addr_l", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_vn.ToString();
                        command.Parameters.Add("p_addr_f", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_kr.ToString();
                        command.Parameters.Add("p_tel", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Phone.ToString();
                        command.Parameters.Add("p_fax", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Fax.ToString();
                        command.Parameters.Add("p_email", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Email.ToString();
                        command.Parameters.Add("p_acc_no", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_no.ToString();
                        command.Parameters.Add("p_acc_ccy", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_ccy.ToString();
                        command.Parameters.Add("p_acc_holder", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_holder.ToString();
                        command.Parameters.Add("p_bank_name", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Bank_name.ToString();
                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tei_company_pk.ToString();
                        command.Parameters.Add("p_remarks", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Remarks.ToString();
                        command.Parameters.Add("p_use_yn", OracleType.VarChar, 1000).Value = "Y";
                        command.Parameters.Add("p_web_site", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Web_site.ToString();
                        command.Parameters.Add("p_erp_customer_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Erp_customer_pk.ToString();
                        command.Parameters.Add("p_tco_buspartner_pk_itf", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Erp_customer_pk.ToString();
                        command.Parameters.Add("p_buyer_name", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Buyer_name.ToString();
                        command.Parameters.Add("p_user_login", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.User_login.ToString();
                        command.Parameters.Add("p_company_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Company_tax_code.ToString();
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Crt_by.ToString();
                        command.Parameters.Add("p_Tax_code_To_UserID", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tax_code_To_UserID.ToString();
                        command.Parameters.Add("p_rtn_master_pk", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;

                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();

                        result = command.Parameters["p_rtn_master_pk"].Value.ToString();
                        connection.Close();
                        connection.Dispose();
                        //return s;


                    }

                }

                return "{\"result\":" + result + ",\"msg\":\"OK\"}";



            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }

        [WebMethod]
        public string UpdateCustomers_v2(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string result = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {

                XmlSerializer serializer = new XmlSerializer(typeof(Customers));

                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Customers _Customers = (Customers)serializer.Deserialize(reader);
                    if (_Customers == null)
                    {// chuyen toi login page

                    }
                    else
                    {

                        string Procedure = "sp_upd_agic00050_upd";
                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "UPDATE";
                        command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_cus_cd", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_cd.ToString();
                        command.Parameters.Add("p_cus_nm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_nm.ToString();
                        command.Parameters.Add("p_cus_lnm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_lnm.ToString();
                        command.Parameters.Add("p_cus_fnm", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Cus_fnm.ToString();
                        command.Parameters.Add("p_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tax_code.ToString();
                        command.Parameters.Add("p_addr", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_en.ToString();
                        command.Parameters.Add("p_addr_l", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_vn.ToString();
                        command.Parameters.Add("p_addr_f", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Address_kr.ToString();
                        command.Parameters.Add("p_tel", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Phone.ToString();
                        command.Parameters.Add("p_fax", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Fax.ToString();
                        command.Parameters.Add("p_email", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Email.ToString();
                        command.Parameters.Add("p_acc_no", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_no.ToString();
                        command.Parameters.Add("p_acc_ccy", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_ccy.ToString();
                        command.Parameters.Add("p_acc_holder", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Acc_holder.ToString();
                        command.Parameters.Add("p_bank_name", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Bank_name.ToString();
                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tei_company_pk.ToString();
                        command.Parameters.Add("p_remarks", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Remarks.ToString();
                        command.Parameters.Add("p_use_yn", OracleType.VarChar, 1000).Value = "Y";
                        command.Parameters.Add("p_web_site", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Web_site.ToString();
                        command.Parameters.Add("p_erp_customer_pk", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Erp_customer_pk.ToString();
                        command.Parameters.Add("p_tco_buspartner_pk_itf", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Erp_customer_pk.ToString();
                        command.Parameters.Add("p_buyer_name", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Buyer_name.ToString();
                        command.Parameters.Add("p_user_login", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.User_login.ToString();
                        command.Parameters.Add("p_company_tax_code", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Company_tax_code.ToString();
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Crt_by.ToString();
                        command.Parameters.Add("p_Tax_code_To_UserID", OracleType.VarChar, 1000).Value = _Customers.Cus[0].Customer.Tax_code_To_UserID.ToString();
                        command.Parameters.Add("p_rtn_master_pk", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;

                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();

                        result = command.Parameters["p_rtn_master_pk"].Value.ToString();
                        connection.Close();
                        connection.Dispose();
                        //return s;


                    }

                }

                return "{\"result\":" + result + ",\"msg\":\"OK\"}";



            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }


        [WebMethod]
        public string Test(string arg_XmlStr)
        {
            return arg_XmlStr;
        }


        [WebMethod]
        public string AdjustInvoice(string arg_XmlStr)
        {
            //  ESysLib.WriteLogError(arg_XmlStr);
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string p_tei_einvoice_m_pk_out = "", p_invoice_no_rtn = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                    if (_Invoices != null)
                    {



                        //  temp = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        string Procedure = "sp_upd_60110350_adjust"; // sp_upd_60110350_adjust_crac

                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                        command.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_company_pk;
                        command.Parameters.Add("p_company_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_name.ToString();
                        command.Parameters.Add("p_company_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_tax_code.ToString();
                        command.Parameters.Add("p_customer_cd", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        command.Parameters.Add("p_customer_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusName.ToString();
                        command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_customer_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Account_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk_vat", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_abacctcode_pk_vat.ToString();
                        command.Parameters.Add("p_tr_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.TrsDate.ToString();
                        command.Parameters.Add("p_tr_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_serial_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Serial_no.ToString();
                        command.Parameters.Add("p_invoice_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoicedate.ToString();
                        command.Parameters.Add("p_form_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Form_No.ToString();
                        command.Parameters.Add("p_invoice_no", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_tr_ccy", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_ccy.ToString();
                        command.Parameters.Add("p_tr_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Trs_exrate.ToString();
                        command.Parameters.Add("p_tr_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_type.ToString();
                        command.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Description.ToString();
                        command.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Local_Description.ToString();
                        command.Parameters.Add("p_remark3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Remark3.ToString();
                        command.Parameters.Add("p_tot_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total_Trans;
                        command.Parameters.Add("p_tot_net_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total;
                        command.Parameters.Add("p_tot_net_tax_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total;
                        command.Parameters.Add("p_tot_vat_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount_Trans;
                        command.Parameters.Add("p_tot_vat_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount;
                        command.Parameters.Add("p_vat_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATRate.ToString();
                        command.Parameters.Add("p_pay_method", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.PaymentMethod.ToString();
                        command.Parameters.Add("p_printed_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_cancel_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_sample_yn", OracleType.VarChar, 1000).Value = "N";
                        command.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk.ToString();
                        command.Parameters.Add("p_invoice_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Type.ToString();
                        command.Parameters.Add("p_ei_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_addr1", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress1;
                        command.Parameters.Add("p_addr2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress2;
                        command.Parameters.Add("p_addr3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress3;
                        command.Parameters.Add("p_bk_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Book_exrate;
                        command.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn;
                        command.Parameters.Add("p_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusTaxCode;
                        command.Parameters.Add("p_invoice_desc", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc;
                        command.Parameters.Add("p_invoice_desc2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc2;
                        command.Parameters.Add("p_DECLARE_NO", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclNo;
                        command.Parameters.Add("p_DECLARE_DATE", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclDate;
                        command.Parameters.Add("p_Adjust_Type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.AdjustType;
                        //         command.Parameters.Add("p_After", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.AfterAmount;InvoicePkRef
                        //        command.Parameters.Add("p_Befor", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.BeforAmount;
                        //        command.Parameters.Add("p_Gap_Amount", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.GapAmount;
                        command.Parameters.Add("p_Invoice_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.InvoiceRef;
                        command.Parameters.Add("p_Invoice_Ref_Dt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.InvoiceRef_Dt;

                        command.Parameters.Add("p_Serial_No_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.SerialNoRef;
                        command.Parameters.Add("p_Form_No_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.FormNoRef;
                        command.Parameters.Add("p_Invoice_Pk_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.InvoicePkRef;
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id;

                        command.Parameters.Add("p_tei_einvoice_m_pk_out", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;
                        command.Parameters.Add("p_invoice_no_rtn", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;


                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();
                        p_tei_einvoice_m_pk_out = command.Parameters["p_tei_einvoice_m_pk_out"].Value.ToString();
                        p_invoice_no_rtn = command.Parameters["p_invoice_no_rtn"].Value.ToString();

                        //  string remark_adjust = "";
                        //  string l_year = _Invoices.Inv[0].Invoice.InvoiceRef_Dt.Substring(0,4);
                        //  string l_month= _Invoices.Inv[0].Invoice.InvoiceRef_Dt.Substring(4,2);
                        //  string l_day = _Invoices.Inv[0].Invoice.InvoiceRef_Dt.Substring(6,2);

                        // if (_Invoices.Inv[0].Invoice.AdjustType =="2")
                        //  {
                        //      remark_adjust = "Điều chỉnh tăng cho hóa đơn số "+ _Invoices.Inv[0].Invoice.InvoiceRef + " ,Ký hiệu hóa đơn:"+ _Invoices.Inv[0].Invoice.SerialNoRef + ", Ngày "+l_day+ " Tháng "+ l_month+ " Năm "+ l_year;
                        // }
                        // else
                        //  {
                        //      remark_adjust = "Điều chỉnh giảm cho hóa đơn số " + _Invoices.Inv[0].Invoice.InvoiceRef + " ,Ký hiệu hóa đơn:" + _Invoices.Inv[0].Invoice.SerialNoRef + ", Ngày " + l_day + " Tháng " + l_month + " Năm " + l_year;
                        //  }
                        for (int i = 0; i < _Invoices.Inv[0].Invoice.Products.Product.Count + 1; i++)
                        {

                            Procedure = "sp_upd_60110350_adjust_d_v2"; //sp_upd_60110350_adjust_d_v2 
                            OracleCommand command_d = new OracleCommand(Procedure, connection);
                            command_d.CommandType = CommandType.StoredProcedure;


                            if (i == _Invoices.Inv[0].Invoice.Products.Product.Count)
                            {

                                command_d.CommandType = CommandType.StoredProcedure;
                                command_d.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                                command_d.Parameters.Add("p_tei_einvoice_d_pk", OracleType.VarChar, 1000).Value = 0;
                                command_d.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = p_tei_einvoice_m_pk_out;
                                command_d.Parameters.Add("p_tco_item_pk", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_qty", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_u_price", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_net_tr_amt", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_tac_crcad_pk", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_item_uom", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id.ToString();
                                command_d.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[0].tac_crcad_pk.ToString();
                                command_d.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn.ToString();
                                command_d.Parameters.Add("p_item_code", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_item_name", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_OrderNo", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[0].OrderNo.ToString();
                                command_d.Parameters.Add("p_Seq", OracleType.VarChar, 1000).Value = i + 1;
                                command_d.Parameters.Add("p_Seq_Dis", OracleType.VarChar, 1000).Value = "";

                                command_d.Parameters.Add("p_Att_01", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_Att_02", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.AdjustInfo;
                                command_d.Parameters.Add("p_Att_03", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_Att_04", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_Att_05", OracleType.VarChar, 1000).Value = "";
                            }
                            else
                            {

                                command_d.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                                command_d.Parameters.Add("p_tei_einvoice_d_pk", OracleType.VarChar, 1000).Value = 0;
                                command_d.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = p_tei_einvoice_m_pk_out;
                                command_d.Parameters.Add("p_tco_item_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tco_item_pk.ToString();
                                command_d.Parameters.Add("p_qty", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdQuantity.ToString();
                                command_d.Parameters.Add("p_u_price", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdPrice.ToString();
                                command_d.Parameters.Add("p_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Amount.ToString();
                                command_d.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DDescription.ToString();
                                command_d.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DLDescription.ToString();
                                command_d.Parameters.Add("p_tac_crcad_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk.ToString();
                                command_d.Parameters.Add("p_item_uom", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdUnit.ToString();
                                command_d.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id.ToString();
                                command_d.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tac_crcad_pk.ToString();
                                command_d.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn.ToString();
                                command_d.Parameters.Add("p_item_code", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Code.ToString();
                                command_d.Parameters.Add("p_item_name", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdName.ToString();
                                command_d.Parameters.Add("p_OrderNo", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].OrderNo.ToString();
                                command_d.Parameters.Add("p_Seq", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Seq.ToString();
                                command_d.Parameters.Add("p_Seq_Dis", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Seq_Dis.ToString();

                                command_d.Parameters.Add("p_Att_01", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_01.ToString();
                                command_d.Parameters.Add("p_Att_02", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_02.ToString();
                                command_d.Parameters.Add("p_Att_03", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_03.ToString();
                                command_d.Parameters.Add("p_Att_04", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_04.ToString();
                                command_d.Parameters.Add("p_Att_05", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_05.ToString();

                            }

                            OracleDataAdapter da_d = new OracleDataAdapter(command_d);
                            command_d.ExecuteNonQuery();

                        }
                        connection.Close();
                        connection.Dispose();

                    }

                }
                return "{\"result\":" + p_tei_einvoice_m_pk_out + "," + "\"Invoice_no\":\"" + p_invoice_no_rtn + "\",\"msg\":\"OK\"}";
            }
            catch (Exception ex)
            {
                //     ESysLib.WriteLogError(arg_XmlStr);
                return ex.Message;
            }
        }

        [WebMethod]
        public string ReplaceInvoice(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string p_tei_einvoice_m_pk_out = "", p_invoice_no_rtn = "", p_rtn_value1 = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Invoices _Invoices = (Invoices)serializer.Deserialize(reader);
                    if (_Invoices != null)
                    {
                        //  temp = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        string Procedure = "sp_upd_60110360_replace"; // sp_upd_60110350_adjust_crac

                        connection = new OracleConnection(_conString);
                        connection.Open();
                        OracleCommand command = new OracleCommand(Procedure, connection);
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                        command.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = 0;
                        command.Parameters.Add("p_tei_company_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_company_pk;
                        command.Parameters.Add("p_company_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_name.ToString();
                        command.Parameters.Add("p_company_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.company_tax_code.ToString();
                        command.Parameters.Add("p_customer_cd", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusCode.ToString();
                        command.Parameters.Add("p_customer_nm", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusName.ToString();
                        command.Parameters.Add("p_tei_customer_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tei_customer_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Account_pk.ToString();
                        command.Parameters.Add("p_tac_abacctcode_pk_vat", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_abacctcode_pk_vat.ToString();
                        command.Parameters.Add("p_tr_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.TrsDate.ToString();
                        command.Parameters.Add("p_tr_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_serial_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Serial_no.ToString();
                        command.Parameters.Add("p_invoice_date", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoicedate.ToString();
                        command.Parameters.Add("p_form_no", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Form_No.ToString();
                        command.Parameters.Add("p_invoice_no", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_tr_ccy", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_ccy.ToString();
                        command.Parameters.Add("p_tr_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Trs_exrate.ToString();
                        command.Parameters.Add("p_tr_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tr_type.ToString();
                        command.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Description.ToString();
                        command.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Local_Description.ToString();
                        command.Parameters.Add("p_remark3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Remark3.ToString();
                        command.Parameters.Add("p_tot_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total_Trans;
                        command.Parameters.Add("p_tot_net_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total;
                        command.Parameters.Add("p_tot_net_tax_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Total;
                        command.Parameters.Add("p_tot_vat_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount_Trans;
                        command.Parameters.Add("p_tot_vat_bk_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATAmount;
                        command.Parameters.Add("p_vat_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.VATRate.ToString();
                        command.Parameters.Add("p_pay_method", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.PaymentMethod.ToString();
                        command.Parameters.Add("p_printed_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_cancel_yn", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_sample_yn", OracleType.VarChar, 1000).Value = "N";
                        command.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk.ToString();
                        command.Parameters.Add("p_invoice_type", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Type.ToString();
                        command.Parameters.Add("p_ei_status", OracleType.VarChar, 1000).Value = "";
                        command.Parameters.Add("p_addr1", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress1;
                        command.Parameters.Add("p_addr2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress2;
                        command.Parameters.Add("p_addr3", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusAddress3;
                        command.Parameters.Add("p_bk_rate", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Book_exrate;
                        command.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn;
                        command.Parameters.Add("p_taxcode", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.CusTaxCode;
                        command.Parameters.Add("p_invoice_desc", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc;
                        command.Parameters.Add("p_invoice_desc2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Invoice_Desc2;
                        command.Parameters.Add("p_DECLARE_NO", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclNo;
                        command.Parameters.Add("p_DECLARE_DATE", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.DeclDate;
                        command.Parameters.Add("p_Invoice_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.InvoiceRef;
                        command.Parameters.Add("p_Invoice_Ref_Dt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.InvoiceRef_Dt;
                        command.Parameters.Add("p_Serial_No_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.SerialNoRef;
                        command.Parameters.Add("p_Form_No_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.FormNoRef;
                        command.Parameters.Add("p_Invoice_Pk_Ref", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.InvoicePkRef;
                        command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id;
                        command.Parameters.Add("p_tei_einvoice_m_pk_out", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;
                        command.Parameters.Add("p_invoice_no_rtn", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;


                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();
                        p_tei_einvoice_m_pk_out = command.Parameters["p_tei_einvoice_m_pk_out"].Value.ToString();
                        p_invoice_no_rtn = command.Parameters["p_invoice_no_rtn"].Value.ToString();

                        //  string remark_adjust = "";
                        //  string l_year = _Invoices.Inv[0].Invoice.InvoiceRef_Dt.Substring(0, 4);
                        //  string l_month = _Invoices.Inv[0].Invoice.InvoiceRef_Dt.Substring(4, 2);
                        //  string l_day = _Invoices.Inv[0].Invoice.InvoiceRef_Dt.Substring(6, 2);

                        //    remark_adjust = "Thay thế cho hóa đơn số " + _Invoices.Inv[0].Invoice.InvoiceRef + " ,Ký hiệu hóa đơn:" + _Invoices.Inv[0].Invoice.SerialNoRef + ", Ngày " + l_day + " Tháng " + l_month + " Năm " + l_year;
                        int count_details = 0;
                        if (_Invoices.Inv[0].Invoice.ReplaceInfo == "")
                        {
                            count_details = _Invoices.Inv[0].Invoice.Products.Product.Count;
                        }
                        else
                        {
                            count_details = _Invoices.Inv[0].Invoice.Products.Product.Count + 1;
                        }
                        for (int i = 0; i < count_details; i++)
                        {

                            Procedure = "sp_upd_60110360_replace_d"; //sp_upd_60110350_adjust_d_v2 
                            OracleCommand command_d = new OracleCommand(Procedure, connection);
                            command_d.CommandType = CommandType.StoredProcedure;


                            if (i == _Invoices.Inv[0].Invoice.Products.Product.Count && _Invoices.Inv[0].Invoice.ReplaceInfo != "")
                            {
                                command_d.CommandType = CommandType.StoredProcedure;
                                command_d.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                                command_d.Parameters.Add("p_tei_einvoice_d_pk", OracleType.VarChar, 1000).Value = 0;
                                command_d.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = p_tei_einvoice_m_pk_out;
                                command_d.Parameters.Add("p_tco_item_pk", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_qty", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_u_price", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_net_tr_amt", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_tac_crcad_pk", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_item_uom", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id.ToString();
                                command_d.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[0].tac_crcad_pk.ToString();
                                command_d.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn.ToString();
                                command_d.Parameters.Add("p_item_code", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_item_name", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_OrderNo", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[0].OrderNo.ToString();
                                command_d.Parameters.Add("p_Seq", OracleType.VarChar, 1000).Value = i + 1;
                                command_d.Parameters.Add("p_Seq_Dis", OracleType.VarChar, 1000).Value = "";

                                command_d.Parameters.Add("p_Att_01", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_Att_02", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.ReplaceInfo;
                                command_d.Parameters.Add("p_Att_03", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_Att_04", OracleType.VarChar, 1000).Value = "";
                                command_d.Parameters.Add("p_Att_05", OracleType.VarChar, 1000).Value = "";
                            }
                            else
                            {

                                command_d.Parameters.Add("p_action", OracleType.VarChar, 1000).Value = "INSERT";
                                command_d.Parameters.Add("p_tei_einvoice_d_pk", OracleType.VarChar, 1000).Value = 0;
                                command_d.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = p_tei_einvoice_m_pk_out;
                                command_d.Parameters.Add("p_tco_item_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tco_item_pk.ToString();
                                command_d.Parameters.Add("p_qty", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdQuantity.ToString();
                                command_d.Parameters.Add("p_u_price", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdPrice.ToString();
                                command_d.Parameters.Add("p_net_tr_amt", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Amount.ToString();
                                command_d.Parameters.Add("p_remark", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DDescription.ToString();
                                command_d.Parameters.Add("p_remark2", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].DLDescription.ToString();
                                command_d.Parameters.Add("p_tac_crcad_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.tac_crca_pk.ToString();
                                command_d.Parameters.Add("p_item_uom", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdUnit.ToString();
                                command_d.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.User_id.ToString();
                                command_d.Parameters.Add("p_tac_crca_pk", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].tac_crcad_pk.ToString();
                                command_d.Parameters.Add("p_interface_itemcode_yn", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.interface_itemcode_yn.ToString();
                                command_d.Parameters.Add("p_item_code", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Code.ToString();
                                command_d.Parameters.Add("p_item_name", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].ProdName.ToString();
                                command_d.Parameters.Add("p_OrderNo", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].OrderNo.ToString();
                                command_d.Parameters.Add("p_Seq", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Seq.ToString();
                                command_d.Parameters.Add("p_Seq_Dis", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Seq_Dis.ToString();

                                command_d.Parameters.Add("p_Att_01", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_01.ToString();
                                command_d.Parameters.Add("p_Att_02", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_02.ToString();
                                command_d.Parameters.Add("p_Att_03", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_03.ToString();
                                command_d.Parameters.Add("p_Att_04", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_04.ToString();
                                command_d.Parameters.Add("p_Att_05", OracleType.VarChar, 1000).Value = _Invoices.Inv[0].Invoice.Products.Product[i].Attribute_05.ToString();

                            }

                            OracleDataAdapter da_d = new OracleDataAdapter(command_d);
                            command_d.ExecuteNonQuery();

                        }
                        connection.Close();
                        connection.Dispose();

                    }

                }
                return "{\"result\":" + p_tei_einvoice_m_pk_out + "," + "\"Invoice_no\":\"" + p_invoice_no_rtn + "\",\"msg\":\"OK\"}";
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        [WebMethod]
        public string Get_Count_EInvoice(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            int result = 0; string xml_json = "";
            OracleConnection connection;


            char[] sep = { '|' };
            string[] param = arg_XmlStr.Split(sep);


            // Display the sentences.

            try
            {
                string Procedure = "sp_sel_Count_Invoice";

                connection = new OracleConnection(_conString);
                connection.Open();
                OracleCommand command = new OracleCommand(Procedure, connection);
                command.CommandType = CommandType.StoredProcedure;


                command.Parameters.Add("p_tco_company_pk", OracleType.VarChar, 1000).Value = param[0].ToString();
                command.Parameters.Add("p_dtfrm", OracleType.VarChar, 1000).Value = param[1].ToString();
                command.Parameters.Add("p_dtto", OracleType.VarChar, 1000).Value = param[2].ToString();
                command.Parameters.Add("p_serial_no", OracleType.VarChar, 1000).Value = param[3].ToString();
                command.Parameters.Add("p_status", OracleType.VarChar, 1000).Value = param[4].ToString();
                command.Parameters.Add("p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
                DataSet ds = new DataSet();
                OracleDataAdapter da = new OracleDataAdapter(command);
                da.Fill(ds);
                DataTable dt = ds.Tables[0];
                result = dt.Rows.Count;

                connection.Close();
                connection.Dispose();

                var JSONString = new StringBuilder();
                //if (dt.Rows.Count > 0)
                //{
                JSONString.Append("[");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    JSONString.Append("{");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j < dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\",");
                        }
                        else if (j == dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\"");
                        }
                    }
                    if (i == dt.Rows.Count - 1)
                    {
                        JSONString.Append("}");
                    }
                    else
                    {
                        JSONString.Append("},");
                    }
                }
                JSONString.Append("]");
                //}
                xml_json = JSONString.ToString();
                return xml_json;
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        [WebMethod]
        public string Send_Bill_Erp(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            int result = 0; string xml_json = "";
            OracleConnection connection;

            char[] sep = { '|' };
            string[] param = arg_XmlStr.Split(sep);


            try
            {
                string Procedure = "AC_PRO_GETBILL_ERP";

                connection = new OracleConnection(_conString);
                connection.Open();
                OracleCommand command = new OracleCommand(Procedure, connection);
                command.CommandType = CommandType.StoredProcedure;


                command.Parameters.Add("p_tei_company", OracleType.VarChar, 1000).Value = param[0].ToString();
                command.Parameters.Add("p_store_id", OracleType.VarChar, 1000).Value = param[1].ToString();
                command.Parameters.Add("p_dtfrm", OracleType.VarChar, 1000).Value = param[2].ToString();
                command.Parameters.Add("p_dtto", OracleType.VarChar, 1000).Value = param[3].ToString();
                command.Parameters.Add("p_crt_by", OracleType.VarChar, 1000).Value = param[4].ToString();
                command.Parameters.Add("p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
                DataSet ds = new DataSet();
                OracleDataAdapter da = new OracleDataAdapter(command);
                da.Fill(ds);
                DataTable dt = ds.Tables[0];
                result = dt.Rows.Count;

                connection.Close();
                connection.Dispose();

                var JSONString = new StringBuilder();
                //if (dt.Rows.Count > 0)
                //{
                JSONString.Append("[");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    JSONString.Append("{");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j < dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\",");
                        }
                        else if (j == dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\"");
                        }
                    }
                    if (i == dt.Rows.Count - 1)
                    {
                        JSONString.Append("}");
                    }
                    else
                    {
                        JSONString.Append("},");
                    }
                }
                JSONString.Append("]");
                //}
                xml_json = JSONString.ToString();
                return xml_json;
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        [WebMethod]
        public string Get_File_Xml(string arg_XmlStr)
        {
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            int result = 0; string xml_json = "";
            OracleConnection connection;


            char[] sep = { '|' };
            string[] param = arg_XmlStr.Split(sep);


            // Display the sentences.

            try
            {
                string Procedure = "sp_sel_Count_Invoice";

                connection = new OracleConnection(_conString);
                connection.Open();
                OracleCommand command = new OracleCommand(Procedure, connection);
                command.CommandType = CommandType.StoredProcedure;


                command.Parameters.Add("p_tco_company_pk", OracleType.VarChar, 1000).Value = param[0].ToString();
                command.Parameters.Add("p_dtfrm", OracleType.VarChar, 1000).Value = param[1].ToString();
                command.Parameters.Add("p_dtto", OracleType.VarChar, 1000).Value = param[2].ToString();
                command.Parameters.Add("p_serial_no", OracleType.VarChar, 1000).Value = param[3].ToString();
                command.Parameters.Add("p_status", OracleType.VarChar, 1000).Value = param[4].ToString();
                command.Parameters.Add("p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
                DataSet ds = new DataSet();
                OracleDataAdapter da = new OracleDataAdapter(command);
                da.Fill(ds);
                DataTable dt = ds.Tables[0];
                result = dt.Rows.Count;

                connection.Close();
                connection.Dispose();

                var JSONString = new StringBuilder();
                //if (dt.Rows.Count > 0)
                //{
                JSONString.Append("[");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    JSONString.Append("{");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j < dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\",");
                        }
                        else if (j == dt.Columns.Count - 1)
                        {
                            JSONString.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":" + "\"" + dt.Rows[i][j].ToString() + "\"");
                        }
                    }
                    if (i == dt.Rows.Count - 1)
                    {
                        JSONString.Append("}");
                    }
                    else
                    {
                        JSONString.Append("},");
                    }
                }
                JSONString.Append("]");
                //}
                xml_json = JSONString.ToString();
                return xml_json;
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }

        private void CreatePO(string filename,int stt)
        {
            // Creates an instance of the XmlSerializer class;
            // specifies the type of object to serialize.
            XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
            TextWriter writer = new StreamWriter(filename);
            Invoices inv = new Invoices();

            // Creates an address to ship and bill to.
           // inv.Inv[stt];
            serializer.Serialize(writer, inv.Inv[stt]);
            writer.Close();
        }

        protected void ReadPO(string filename)
        {
            // Creates an instance of the XmlSerializer class;
            // specifies the type of object to be deserialized.
            XmlSerializer serializer = new XmlSerializer(typeof(Invoices));
            // If the XML document has been altered with unknown
            // nodes or attributes, handles them with the
            // UnknownNode and UnknownAttribute events.
           /* serializer.UnknownNode += new
            XmlNodeEventHandler(serializer_UnknownNode);
            serializer.UnknownAttribute += new
            XmlAttributeEventHandler(serializer_UnknownAttribute);

            // A FileStream is needed to read the XML document.
            FileStream fs = new FileStream(filename, FileMode.Open);
            // Declares an object variable of the type to be deserialized.
            PurchaseOrder po;
            // Uses the Deserialize method to restore the object's state
            // with data from the XML document. 
            po = (PurchaseOrder)serializer.Deserialize(fs);
            // Reads the order date.
            Console.WriteLine("OrderDate: " + po.OrderDate);

            // Reads the shipping address.
            Address shipTo = po.ShipTo;
            ReadAddress(shipTo, "Ship To:");
            // Reads the list of ordered items.
            OrderedItem[] items = po.OrderedItems;
            Console.WriteLine("Items to be shipped:");
            foreach (OrderedItem oi in items)
            {
                Console.WriteLine("\t" +
                oi.ItemName + "\t" +
                oi.Description + "\t" +
                oi.UnitPrice + "\t" +
                oi.Quantity + "\t" +
                oi.LineTotal);
            }
            // Reads the subtotal, shipping cost, and total cost.
            Console.WriteLine(
            "\n\t\t\t\t\t Subtotal\t" + po.SubTotal +
            "\n\t\t\t\t\t Shipping\t" + po.ShipCost +
            "\n\t\t\t\t\t Total\t\t" + po.TotalCost
            );*/
        }
    }

}
