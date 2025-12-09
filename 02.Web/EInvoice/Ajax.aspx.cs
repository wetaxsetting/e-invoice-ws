using System;
using System.Web.Services;
using System.Xml.Serialization;
using System.Data;
using System.Data.OracleClient;
using System.IO;
namespace EInvoice
{
    public partial class Ajax : System.Web.UI.Page
    {
        private string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
        // string connString = String.Format( "Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
        private string dbName = "STDGW_242_HAMACHI", dbUser = "stdgw", dbPwd = "stdgw2";
        protected void Page_Load(object sender, EventArgs e)
        {
            string arg_XmlStr = Request["arg_XmlStr"];

            UpdateCustomers(arg_XmlStr);
        }

        public void UpdateCustomers(string arg_XmlStr)
        {
            Response.ContentType = "application/json";
            Response.Charset = "utf-8";
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);
            string result = "";
            // connString = String.Format("Data Source={0};Password={1};User ID={2};", m_Host, m_User, m_Password);
            OracleConnection connection;
        

                XmlSerializer serializer = new XmlSerializer(typeof(Customers));

                using (TextReader reader = new StringReader(arg_XmlStr))
                {
                    Customers _Customers = (Customers)serializer.Deserialize(reader);
                    if (_Customers == null)
                    {// chuyen toi login page

                    }
                    else
                    {

                        string exeStatement = "", temp = "";
                        string Procedure = "sp_upd_agic00050_3";
                        int i = 0;
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
                        command.Parameters.Add("p_rtn_master_pk", OracleType.VarChar, 1000).Direction = ParameterDirection.Output;
                        /*
                        command.Parameters.Add(new OracleParameter(":p_cus_cd", _Customers.Customer[0].Cus_cd.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_cus_nm", _Customers.Customer[0].Cus_nm.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_cus_lnm", _Customers.Customer[0].Cus_lnm.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_cus_fnm", _Customers.Customer[0].Cus_fnm.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_tax_code", _Customers.Customer[0].Tax_code.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_addr", _Customers.Customer[0].Address_en.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_addr_l", _Customers.Customer[0].Address_vn.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_addr_f", _Customers.Customer[0].Address_kr.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_tel", _Customers.Customer[0].Phone.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_fax", _Customers.Customer[0].Fax.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_email", _Customers.Customer[0].Email.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_acc_no", _Customers.Customer[0].Acc_no.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_acc_ccy", _Customers.Customer[0].Acc_ccy.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_acc_holder", _Customers.Customer[0].Acc_holder.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_bank_name", _Customers.Customer[0].Bank_name.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_tei_company_pk", _Customers.Customer[0].Tei_company_pk.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_remarks", _Customers.Customer[0].Remarks.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_use_yn", "Y"));
                        command.Parameters.Add(new OracleParameter(":p_web_site", _Customers.Customer[0].Web_site.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_erp_customer_pk", _Customers.Customer[0].Erp_customer_pk.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_tco_buspartner_pk_itf", _Customers.Customer[0].Erp_customer_pk.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_buyer_name", _Customers.Customer[0].Buyer_name.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_user_login", _Customers.Customer[0].User_login.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_company_tax_code", _Customers.Customer[0].Company_tax_code.ToString()));
                        command.Parameters.Add(new OracleParameter(":p_crt_by", _Customers.Customer[0].Crt_by.ToString()));

                        exeStatement = "Call " + Procedure + "(" + p_parameter01 + "," + p_parameter02 + ",:p_rtn_value)";
                        */
                        // command.Parameters.Add[":p_rtn_master_pk"].Direction = ParameterDirection.Output;
                        // command.Parameters.Add("@p_rtn_master_pk", OracleType.Number).Direction = ParameterDirection.Output;
                        //objCmd.Parameters.Add("pout_count", OracleType.Number).Direction = ParameterDirection.Output;
                        // command.Parameters[":p_rtn_master_pk"].Direction = ParameterDirection.Output;
                        // exeStatement = "call Procedure ";


                        OracleDataAdapter da = new OracleDataAdapter(command);
                        command.ExecuteNonQuery();

                        result = command.Parameters["p_rtn_master_pk"].Value.ToString();
                        connection.Close();
                        connection.Dispose();
                        //return s;


                    }

                }
                Response.Clear();
                Response.Write("{\"result\":"+ result + ",\"msg\":\"OK\"}");
                Response.End();


        }
    }
}