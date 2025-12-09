using System;
using System.Data;
//using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;
using System.Collections.Generic;

using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;


namespace EInvoice.Company
{
    public class CucKoo_C23TCB_New
    {
        public static string View(string tei_einvoice_m_pk, string tei_company_pk, string dbName)
        {
            /*string dbUser = "genuwin", dbPwd = "genuwin2";//NOBLANDBD  EINVOICE_252
            string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);*/
            string _conString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=123.30.104.243)(PORT=1941))(CONNECT_DATA=(SERVER=dedicated)(SERVICE_NAME=NOBLANDBD)));User ID=genuwin;Password=genuwin2";


            string Procedure = "stacfdstac71_r_02_view_einv_v2"; //stacfdstac71_r_02_1
            OracleConnection connection;
            connection = new OracleConnection(_conString);
            connection.Open();
            OracleCommand command = new OracleCommand(Procedure, connection);
            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.Add("p_tei_einvoice_m_pk", OracleDbType.Varchar2, 1000).Value = tei_einvoice_m_pk;
            command.Parameters.Add("p_user_id", OracleDbType.Varchar2, 1000).Value = "genuwin";
            command.Parameters.Add("p_rtn_value", OracleDbType.RefCursor).Direction = ParameterDirection.Output;

            DataSet ds = new DataSet();

            OracleDataAdapter da = new OracleDataAdapter(command);
            da.Fill(ds);
            DataTable dt = ds.Tables[0];

            command.Parameters.Clear();

            Procedure = "stacfdstac71_r_03_view_einv";  //stacfdstac71_r_03

            command = new OracleCommand(Procedure, connection);
            command.CommandType = CommandType.StoredProcedure;


            command.Parameters.Add("p_tei_einvoice_m_pk", OracleDbType.Varchar2, 1000).Value = tei_einvoice_m_pk;
            command.Parameters.Add("p_user_id", OracleDbType.Varchar2, 1000).Value = "genuwin";
            command.Parameters.Add("p_rtn_value", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
            DataSet ds_d = new DataSet();
            OracleDataAdapter da_d = new OracleDataAdapter(command);
            da_d.Fill(ds_d);
            DataTable dt_d = ds_d.Tables[0];


            command.Parameters.Clear();

            Procedure = "stacfdstac71_r_04_view_einv";  //stacfdstac71_r_03

            command = new OracleCommand(Procedure, connection);
            command.CommandType = CommandType.StoredProcedure;


            command.Parameters.Add("p_tei_einvoice_m_pk", OracleDbType.Varchar2, 1000).Value = tei_einvoice_m_pk;
            command.Parameters.Add("p_user_id", OracleDbType.Varchar2, 1000).Value = "genuwin";
            command.Parameters.Add("p_rtn_value", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
            DataSet ds_dvat = new DataSet();
            OracleDataAdapter da_dvat = new OracleDataAdapter(command);
            da_dvat.Fill(ds_dvat);
            DataTable dt_dvat = ds_dvat.Tables[0];

            int pos = 12, pos_lv = 20, v_count = 0, count_page = 0, count_page_v = 0, r = 0, x = 0;

            v_count = dt_d.Rows.Count;  //_Invoices.Inv[0].Invoice.Products.Product.Count();
            int[] page = new int[10] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            int v_index = -1, rowsPerPage = 20;

            if (v_count % pos_lv == 0)
            {
                int n1 = v_count / pos_lv;
                for (int i = 0; i < n1 - 1; i++)
                {
                    page[i] = 20;
                }
                page[n1 - 1] = 19;
                page[n1] = 1;
            }
            else
            {
                int n2 = v_count / pos_lv;
                for (int i = 0; i < n2; i++)
                {
                    page[i] = 20;
                }

                if (v_count % pos_lv > 11)
                {
                    page[n2] = (v_count % pos_lv) - 1;

                    page[n2 + 1] = 1;
                }
                else
                {
                    page[n2] = (v_count % pos_lv);
                }
            }
            int v_countNumberOfPages = 0;
            for (int i = 0; i < page.Length; i++)
            {
                if (page[i] > 0)
                {
                    v_countNumberOfPages++;
                }
            }

            string read_prive = "", read_en = "", read_amount = "", amout_vat = "";

            read_prive = dt.Rows[0]["amount_word_vie"].ToString();


            if (dt.Rows[0]["vatamount_display"].ToString().Trim() == "0")
            {
                amout_vat = "-";
            }
            else
            {
                amout_vat = dt.Rows[0]["vatamount_display"].ToString();
            }

            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";

            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>												                                                           \n");
            htmlStr.Append("<html>                                                                                                                                                                                                               \n");
            htmlStr.Append("<head>                                                                                                                                                                                                               \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                                                                  \n");
            htmlStr.Append("                                                                                                                                                                                                                     \n");
            htmlStr.Append("<script type='text/javascript'                                                                                                                                              \n");
            htmlStr.Append("	src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                                                                 \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                             \n");
            htmlStr.Append("<!-- Normalize or reset CSS with your favorite library -->                                                                                                                  \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("<!-- Load paper.css for happy printing -->                                                                                                                                  \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("<!-- Set page size here: A5, A4 or A3 -->                                                                                                                                   \n");
            htmlStr.Append("<!-- Set also 'landscape' if you need -->                                                                                                                                   \n");
            htmlStr.Append("<style>                                                                                                                                                                     \n");
            htmlStr.Append("@page {                                                                                                                                                                     \n");
            htmlStr.Append("	size: A4;                                                                                                                                                               \n");
            htmlStr.Append("	padding-left: 50px;                                                                                                                                                     \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("</style>                                                                                                                                                                    \n");
            htmlStr.Append("<style>                                                                                                                                                                     \n");
            htmlStr.Append("/*body   { font-family: serif }                                                                                                                                             \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                         \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                         \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                            \n");
            htmlStr.Append("    h2 + p { font-size: 23pt; line-height: 7mm }                                                                                                                            \n");
            htmlStr.Append("    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                                                            \n");
            htmlStr.Append("    li     { font-size: 11pt; line-height: 5mm }                                                                                                                            \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                   \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                           \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                            \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                 \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                        \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                       \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                        \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                           \n");
            htmlStr.Append("body {                                                                                                                                                                      \n");
            htmlStr.Append("	color: blue;                                                                                                                                                            \n");
            htmlStr.Append("	font-size: 100%;                                                                                                                                                        \n");
            htmlStr.Append("	background-image: url('assets/Solution.jpg');                                                                                                                           \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("h1 {                                                                                                                                                                        \n");
            htmlStr.Append("	color: #00FF00;                                                                                                                                                         \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("p {                                                                                                                                                                         \n");
            htmlStr.Append("	color: rgb(0, 0, 255)                                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("headline1 {                                                                                                                                                                 \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                                             \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                           \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                          \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                      \n");
            htmlStr.Append("	margin-bottom: 50px;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("headline2 {                                                                                                                                                                 \n");
            htmlStr.Append("	background-image: url(images/newsletter_headline2.gif);                                                                                                                 \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                           \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                          \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                      \n");
            htmlStr.Append("	padding-left: 68px;                                                                                                                                                     \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("            <!-- table {																																						  \n");
            htmlStr.Append("                mso-displayed-decimal-separator: '\\.';                                                                                                                            \n");
            htmlStr.Append("                mso-displayed-thousand-separator: '\\,';                                                                                                                           \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font526694 {                                                                                                                                                         \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font626694 {                                                                                                                                                         \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font726694 {                                                                                                                                                         \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font826694 {                                                                                                                                                         \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font926694 {                                                                                                                                                         \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1026694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1126694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1226694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: #0066CC;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1326694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: Arial, sans-serif;                                                                                                                                   \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1426694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: Arial, sans-serif;                                                                                                                                   \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1526694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1626694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1726694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .font1826694 {                                                                                                                                                        \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl6526694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl6626694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.7pt;                                                                                                                                               \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl6726694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl6826694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.7pt;                                                                                                                                               \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border: .5pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl6926694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7026694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: Arial, sans-serif;                                                                                                                                   \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7126694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl71266941 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7226694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7326694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7426694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7526694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7626694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7726694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.90pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7826694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl7926694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8026694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.90pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8126694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8226694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8326694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8426694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8526694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8626694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8726694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8826694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl8926694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9026694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9126694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9226694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9326694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9426694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt solid #2F75B5;                                                                                                                                  \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9526694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt solid #2F75B5;                                                                                                                                  \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9626694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt solid #2F75B5;                                                                                                                                  \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9726694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl97266941 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9826694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl9926694 {                                                                                                                                                          \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl99266941 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 1.0pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl10926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt solid #2F75B5;                                                                                                                               \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.7pt;                                                                                                                                               \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt solid #2F75B5;                                                                                                                                  \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl11926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl121266941 {                                                                                                                                                        \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl12926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: 1.0pt solid #2F75B5;                                                                                                                                 \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: general;                                                                                                                                              \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl13926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 8.50pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 14.0pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 14.0pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl14926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 8.50pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border: .5pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl156266941 {                                                                                                                                                        \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border: .5pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 8.50pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl15926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt solid #2F75B5;                                                                                                                               \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: 1.0pt solid #2F75B5;                                                                                                                                \n");
            htmlStr.Append("                border-bottom: 1.0pt solid #2F75B5;                                                                                                                               \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #FFC000;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #FFC000;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #FFC000;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl16926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: Arial, sans-serif;                                                                                                                                   \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: Arial, sans-serif;                                                                                                                                   \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: Arial, sans-serif;                                                                                                                                   \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: none;                                                                                                                                                 \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl17926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl18926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: .5pt solid windowtext;                                                                                                                             \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border: .5pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt dotted windowtext;                                                                                                                               \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 10.9pt;                                                                                                                                                 \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl19926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: 1.0pt dotted windowtext;                                                                                                                            \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: #0070C0;                                                                                                                                                   \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 12.1pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: .5pt solid windowtext;                                                                                                                               \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20426694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: .5pt solid windowtext;                                                                                                                              \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20526694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: top;                                                                                                                                              \n");
            htmlStr.Append("                border-top: .5pt solid windowtext;                                                                                                                                \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20626694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 14.0pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: normal;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20726694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt solid #2F75B5;                                                                                                                                  \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20826694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: black;                                                                                                                                                     \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl20926694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 14.35pt;                                                                                                                                               \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                border-top: 1.0pt solid #2F75B5;                                                                                                                                  \n");
            htmlStr.Append("                border-right: none;                                                                                                                                               \n");
            htmlStr.Append("                border-bottom: none;                                                                                                                                              \n");
            htmlStr.Append("                border-left: none;                                                                                                                                                \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl21026694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 14.85pt;                                                                                                                                               \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl21126694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: right;                                                                                                                                                \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl21226694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: red;                                                                                                                                                       \n");
            htmlStr.Append("                font-size: 14.0pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 700;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: italic;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: center;                                                                                                                                               \n");
            htmlStr.Append("                vertical-align: bottom;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                  \n");
            htmlStr.Append("            .xl21326694 {                                                                                                                                                         \n");
            htmlStr.Append("                padding: 0px;                                                                                                                                                     \n");
            htmlStr.Append("                mso-ignore: padding;                                                                                                                                              \n");
            htmlStr.Append("                color: windowtext;                                                                                                                                                \n");
            htmlStr.Append("                font-size: 13.3pt;                                                                                                                                                \n");
            htmlStr.Append("                font-weight: 400;                                                                                                                                                 \n");
            htmlStr.Append("                font-style: normal;                                                                                                                                               \n");
            htmlStr.Append("                text-decoration: none;                                                                                                                                            \n");
            htmlStr.Append("                font-family: 'Times New Roman', serif;                                                                                                                            \n");
            htmlStr.Append("                mso-font-charset: 0;                                                                                                                                              \n");
            htmlStr.Append("                mso-number-format: General;                                                                                                                                       \n");
            htmlStr.Append("                text-align: left;                                                                                                                                                 \n");
            htmlStr.Append("                vertical-align: middle;                                                                                                                                           \n");
            htmlStr.Append("                mso-background-source: auto;                                                                                                                                      \n");
            htmlStr.Append("                mso-pattern: auto;                                                                                                                                                \n");
            htmlStr.Append("                white-space: nowrap;                                                                                                                                              \n");
            htmlStr.Append("            }                                                                                                                                                                     \n");
            htmlStr.Append("            -->																																									  \n"); htmlStr.Append("</style>                                                                                                                                                                    \n");
            htmlStr.Append("</head>                                                                                                                                                                     \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                           \n");

            htmlStr.Append("            <table border=0 cellpadding=0 cellspacing=0 width=795 class=xl6526694																								                                                  	\n");
            htmlStr.Append("                   style='border-collapse:collapse;table-layout:fixed;width:648.0pt'>                                                                                                                                             	\n");
            htmlStr.Append("                <col class=xl6526694 width=12 style='mso-width-source:userset;mso-width-alt:426;width:9.8pt'>                                                                                                                     	\n");
            htmlStr.Append("                <col class=xl6526694 width=38 style='mso-width-source:userset;mso-width-alt:1336;width:30.5pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=81 style='mso-width-source:userset;mso-width-alt:2872;width:66.4pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=48 style='mso-width-source:userset;mso-width-alt:1706;width:39.2pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=41 style='mso-width-source:userset;mso-width-alt:1450;width:33.8pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=62 style='mso-width-source:userset;mso-width-alt:2218;width:51.2pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=49 style='mso-width-source:userset;mso-width-alt:1735;width:40.3pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=56 style='mso-width-source:userset;mso-width-alt:1991;width:45.8pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=34 style='mso-width-source:userset;mso-width-alt:1194;width:27.2pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=42 style='mso-width-source:userset;mso-width-alt:1479;width:33.8pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=78 style='mso-width-source:userset;mso-width-alt:2759;width:63.2pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=56 style='mso-width-source:userset;mso-width-alt:1991;width:45.8pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=70 style='mso-width-source:userset;mso-width-alt:2503;width:57.7pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=19 style='mso-width-source:userset;mso-width-alt:682;width:15.2pt'>                                                                                                                    	\n");
            htmlStr.Append("                <col class=xl6526694 width=42 style='mso-width-source:userset;mso-width-alt:1479;width:33.8pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=55 style='mso-width-source:userset;mso-width-alt:1962;width:44.7pt'>                                                                                                                   	\n");
            htmlStr.Append("                <col class=xl6526694 width=12 style='mso-width-source:userset;mso-width-alt:426;width:9.8pt'>                                                                                                                     	\n");


            v_index = 0;
            string v_titlePageNumber = "";
            double v_spacePerPage = 0;

            string v_rowHeight = "25.5pt"; //"26.5pt";
            string v_rowHeightEmpty = "22.0pt";
            double v_rowHeightNumber = 26.5;

            string v_rowHeightLast = "22.5pt";// "23.5pt";
            double v_rowHeightLastNumber = 22.5;// 23.5;
            string v_rowHeightEmptyLast = "22.5pt"; //"23.5pt";

            bool vlongItemName = false;

            double v_totalHeightLastPage = 243.5;// 258.5;   MAX = 895   HEIGH = 384   detail = 497

            double v_totalHeightPage = 525;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 525;// 540;  Math.Ceiling(a)

                if (v_countNumberOfPages > 1)
                {
                    if (k == 0)
                    {
                        v_titlePageNumber = "Trang 1/" + v_countNumberOfPages.ToString();
                    }
                    else if (k < v_countNumberOfPages - 1)
                    {
                        v_titlePageNumber = "tiep theo trang truoc - Trang " + (k + 1).ToString() + "/ " + v_countNumberOfPages.ToString();
                    }
                    else if (k == v_countNumberOfPages - 1)
                    {
                        v_titlePageNumber = "Tiep theo trang truoc - Trang " + (k + 1).ToString() + "/ " + v_countNumberOfPages.ToString();
                    }
                }

                if (k == v_countNumberOfPages - 1)
                {
                    rowsPerPage = pos;
                }
                else
                {
                    rowsPerPage = pos_lv;
                }


                htmlStr.Append("                <tr class=xl7126694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                   	\n");
                htmlStr.Append("                    <td height=22 class=xl9426694 width=12 style='height:17.4pt;width:9.8pt'>&nbsp;</td>                                                                                                                         	\n");
                htmlStr.Append("                    <td class=xl9526694 width=38 style='width:30.5pt'>&nbsp;</td>                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl9526694 width=81 style='width:66.4pt'>&nbsp;</td>                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl9526694 width=48 style='width:39.2pt'>&nbsp;</td>                                                                                                                                                 	\n");
                htmlStr.Append("                    <td colspan=7 rowspan=3 class=xl20926694 width=362 style='width:271pt'>                                                                                                                                       	\n");
                htmlStr.Append("                        HÓA                                                                                                                                                                                                       	\n");
                htmlStr.Append("                        &#272;&#416;N GIÁ TR&#7882; GIA T&#258;NG                                                                                                                                                                 	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                         	\n");
                htmlStr.Append("                    <td class=xl11726694 width=56 style='width:45.8pt'>&nbsp;</td>                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl11726694 width=70 style='width:57.7pt'>&nbsp;</td>                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=3 rowspan=3 class=xl20726694 width=116 style='width:86pt'>&nbsp;</td>                                                                                                                             	\n");
                htmlStr.Append("                    <td class=xl9626694 width=12 style='width:9.8pt'>&nbsp;</td>                                                                                                                                                  	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                             	\n");
                htmlStr.Append("                <tr class=xl71266941 height=10 style='mso-height-source:userset;height:9.6pt'>                                                                                                                                   	\n");
                htmlStr.Append("                    <td height=10 class=xl97266941 style='height:9.6pt'>&nbsp;</td>                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl7326694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl7326694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl7326694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl9026694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl9026694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl9826694>&nbsp;</td>                                                                                                                                                                               	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                             	\n");
                htmlStr.Append("                <tr class=xl71266941 height=6 style='mso-height-source:userset;height:6.0pt'>                                                                                                                                    	\n");
                htmlStr.Append("                    <td height=6 class=xl97266941 style='height:6.0pt'>&nbsp;</td>                                                                                                                                               	\n");
                htmlStr.Append("                    <td colspan=3 rowspan=3 height=49 class=xl7326694 width=167 style='mso-ignore:colspan-rowspan;height:53.2px;width:181.5px'>                                                                                     	\n");
                htmlStr.Append("                        <![if !vml]><span style='mso-ignore:vglayout'>                                                                                                                                                            	\n");
                htmlStr.Append("                            <table cellpadding=0 cellspacing=0>                                                                                                                                                                   	\n");
                htmlStr.Append("                                <tr>                                                                                                                                                                                              	\n");
                htmlStr.Append("                                    <td width=6 height=2></td>                                                                                                                                                                    	\n");
                htmlStr.Append("                                </tr>                                                                                                                                                                                             	\n");
                htmlStr.Append("                                <tr>                                                                                                                                                                                              	\n");
                htmlStr.Append("                                    <td></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                                    <td>                                                                                                                                                                                          	\n");
                htmlStr.Append("                                        <img width=181.5 height=53.2                                                                                                                                                                  	\n");
                htmlStr.Append("                                             src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/CUCKOO_LOGO_NEW.png'                                                                                                           	\n");
                htmlStr.Append("                                             v:shapes='Picture_x0020_6'>                                                                                                                                                          	\n");
                htmlStr.Append("                                    </td>                                                                                                                                                                                         	\n");
                htmlStr.Append("                                    <td width=11></td>                                                                                                                                                                            	\n");
                htmlStr.Append("                                </tr>                                                                                                                                                                                             	\n");
                htmlStr.Append("                                <tr>                                                                                                                                                                                              	\n");
                htmlStr.Append("                                    <td height=3></td>                                                                                                                                                                            	\n");
                htmlStr.Append("                                </tr>                                                                                                                                                                                             	\n");
                htmlStr.Append("                            </table>                                                                                                                                                                                              	\n");
                htmlStr.Append("                        </span><![endif]><!--[if !mso & vml]><span style='width:124.8pt;height:37.2pt'></span><![endif]-->                                                                                                        	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                         	\n");
                htmlStr.Append("                    <td class=xl9026694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl9026694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl9826694>&nbsp;</td>                                                                                                                                                                               	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                             	\n");
                htmlStr.Append("                <tr class=xl7126694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                   	\n");
                htmlStr.Append("                    <td height=22 class=xl9726694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                              	\n");
                htmlStr.Append("                    <td colspan=7 class=xl21226694>VAT INVOICE</td>                                                                                                                                                               	\n");
                htmlStr.Append("                    <td colspan=2 rowspan=2 class=xl21126694>                                                                                                                                                                     	\n");
                htmlStr.Append("                        Ký hi&#7879;u / <font class='font826694'>Serial</font><font class='font1026694'>: </font>                                                                                                                 	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                         	\n");
                htmlStr.Append("                    <td colspan=3 rowspan=2 class=xl20826694>" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</td>                                                                                                                                  	\n");
                htmlStr.Append("                    <td class=xl9826694>&nbsp;</td>                                                                                                                                                                               	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                             	\n");
                htmlStr.Append("                <tr class=xl7126694 height=21 style='height:14.4pt'>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td height=21 class=xl9726694 style='height:14.4pt'>&nbsp;</td>                                                                                                                                               	\n");
                htmlStr.Append("                    <td colspan=7 class=xl15326694></td>                                                                                                                                                                          	\n");
                htmlStr.Append("                    <td class=xl9826694>&nbsp;</td>                                                                                                                                                                               	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                             	\n");
                htmlStr.Append("                <tr height=23 style='height:17.4pt'>                                                                                                                                                                              	\n");
                htmlStr.Append("                    <td height=23 class=xl9926694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                               	\n");
                htmlStr.Append("                    <td class=xl7226694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl7226694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl7226694></td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td colspan=7 class=xl15226694>                                                                                                                                                                               	\n");
                htmlStr.Append("                        Ngày / <font class='font826694'>Date</font><font class='font526694'><span style='mso-spacerun:yes'> " + dt.Rows[0]["invoiceissueddate_dd"] + "  </span>tháng / </font><font class='font826694'>month</font><font class='font526694'>   	\n");
                htmlStr.Append("                            <span style='mso-spacerun:yes'> " + dt.Rows[0]["invoiceissueddate_mm"] + "  </span>n&#259;m /																																			\n");
                htmlStr.Append("                        </font><font class='font826694'>year</font> " + dt.Rows[0]["invoiceissueddate_yyyy"] + " </br> " + v_titlePageNumber + "                                                                                                                    	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                           	\n");
                htmlStr.Append("                    <td colspan=2 class=xl21126694>                                                                                                                                                                                 	\n");
                htmlStr.Append("                        S&#7889; / <font class='font826694'>Invoice no</font><font class='font1026694'>:<span style='mso-spacerun:yes'> </span></font>                                                                              	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                           	\n");
                htmlStr.Append("                    <td colspan=3 class=xl20626694 width=116 style='width:86pt'>" + dt.Rows[0]["InvoiceNumber"] + "</td>                                                                                                                             	\n");
                htmlStr.Append("                    <td class=xl10026694>&nbsp;</td>                                                                                                                                                                                	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                               	\n");
                htmlStr.Append("                                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr height=6 style='mso-height-source:userset;height:6.0pt'>                                                                                                                                                       	\n");
                htmlStr.Append("                    <td height=6 class=xl99266941 style='height:6.0pt'>&nbsp;</td>                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl8926694>&nbsp;</td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                               	\n");
                htmlStr.Append("                <tr height=6 style='mso-height-source:userset;height:6.0pt'>                                                                                                                                                       	\n");
                htmlStr.Append("                    <td height=6 class=xl99266941 style='height:6.0pt'>&nbsp;</td>                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                               	\n");
                htmlStr.Append("                <tr height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td height=22 class=xl9926694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=3>                                                                                                                                                                                  	\n");
                htmlStr.Append("                        &#272;&#417;n v&#7883; bán hàng / <font class='font826694'>Seller</font><font class='font526694'>:</font>                                                                                                   	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                           	\n");
                htmlStr.Append("                    <td colspan=7 class=xl11826694 width=362 style='width:271pt'>" + dt.Rows[0]["Seller_Name"] + "</td>                                                                                                                             	\n");
                htmlStr.Append("                    <td class=xl11826694 width=56 style='width:45.8pt'></td>                                                                                                                                                        	\n");
                htmlStr.Append("                    <td class=xl11826694 width=70 style='width:57.7pt'></td>                                                                                                                                                        	\n");
                htmlStr.Append("                    <td align=left valign=top>                                                                                                                                                                                      	\n");
                htmlStr.Append("                        <![if !vml]><span style='mso-ignore:vglayout;position:absolute;z-index:2;margin-left:4px;margin-top:2px;width:97px;height:92px'>                                                                            	\n");
                htmlStr.Append("                            <img width=97 height=92                                                                                                                                                                                 	\n");
                htmlStr.Append("                                 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/CUCKOO_QR.png'                                                                                                                               	\n");
                htmlStr.Append("                                 v:shapes='Picture_x0020_3'>                                                                                                                                                                        	\n");
                htmlStr.Append("                        </span><![endif]><span style='mso-ignore:vglayout2'>                                                                                                                                                        	\n");
                htmlStr.Append("                            <table cellpadding=0 cellspacing=0>                                                                                                                                                                     	\n");
                htmlStr.Append("                                <tr>                                                                                                                                                                                                	\n");
                htmlStr.Append("                                    <td height=22 class=xl6626694 width=19 style='height:17.4pt;width:15.2pt'></td>                                                                                                                	\n");
                htmlStr.Append("                                </tr>                                                                                                                                                                                               	\n");
                htmlStr.Append("                            </table>                                                                                                                                                                                                	\n");
                htmlStr.Append("                        </span>                                                                                                                                                                                                     	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                           	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                               	\n");
                htmlStr.Append("                <tr height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td height=22 class=xl9926694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=3>                                                                                                                                                                                  	\n");
                htmlStr.Append("                        Mã s&#7889; thu&#7871; / <font class='font826694'>Tax code</font><font class='font526694'>																														 	\n");
                htmlStr.Append("                            :<span style='mso-spacerun:yes'> </span>                                                                                                                                                                     	\n");
                htmlStr.Append("                        </font>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl11426694 colspan=3>" + dt.Rows[0]["Seller_TaxCode"] + "</td>                                                                                                                                                           	\n");
                htmlStr.Append("                    <td class=xl11926694></td>                                                                                                                                                                                           	\n");
                htmlStr.Append("                    <td class=xl11926694></td>                                                                                                                                                                                           	\n");
                htmlStr.Append("                    <td class=xl11926694></td>                                                                                                                                                                                           	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=22 class=xl9926694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=3>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        &#272;&#7883;a ch&#7881; / <font class='font826694'>Address</font><font class='font526694'>                                                                                                                      	\n");
                htmlStr.Append("                            :<span style='mso-spacerun:yes'> </span>                                                                                                                                                                     	\n");
                htmlStr.Append("                        </font>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=9 rowspan=2 class=xl11626694 width=362 style='width:271pt'>" + dt.Rows[0]["Seller_Address"] + "</td>                                                                                                                        	\n");
                //htmlStr.Append("                    <td class=xl11626694 width=56 style='width:45.8pt'></td>                                                                                                                                                             	\n");
                //htmlStr.Append("                    <td class=xl11626694 width=70 style='width:57.7pt'></td>                                                                                                                                                             	\n");
                htmlStr.Append("                    <td class=xl8726694 width=19 style='width:15.2pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=42 style='width:33.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=55 style='width:44.7pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr height=13 style='mso-height-source:userset;height:10.05pt'>                                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=13 class=xl9926694 style='height:10.05pt'>&nbsp;</td>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl6626694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                //htmlStr.Append("                    <td class=xl11626694 width=56 style='width:45.8pt'></td>                                                                                                                                                             	\n");
                //htmlStr.Append("                    <td class=xl11626694 width=70 style='width:57.7pt'></td>                                                                                                                                                             	\n");
                htmlStr.Append("                    <td class=xl8726694 width=19 style='width:15.2pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=42 style='width:33.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=55 style='width:44.7pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr height=21 style='height:17.4pt'>                                                                                                                                                                                    	\n");
                htmlStr.Append("                    <td height=21 class=xl9926694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl7426694 colspan=2>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        &#272;i&#7879;n tho&#7841;i / <font class='font1826694'>Tel:</font><font class='font1726694'>                                                                                                                    	\n");
                htmlStr.Append("                            <span style='mso-spacerun:yes'> </span>                                                                                                                                                                      	\n");
                htmlStr.Append("                        </font>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td colspan=3 class=xl8826694 width=152 style='width:115pt'>" + dt.Rows[0]["Seller_Tel"] + "</td>                                                                                                                                  	\n");
                htmlStr.Append("                    <td class=xl8826694 width=56 style='width:45.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8826694 width=34 style='width:27.2pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8826694 width=42 style='width:33.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8826694 width=78 style='width:63.2pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8826694 width=56 style='width:45.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8826694 width=70 style='width:57.7pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=19 style='width:15.2pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=42 style='width:33.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=55 style='width:44.7pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr height=18 style='mso-height-source:userset;height:13.95pt'>                                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=18 class=xl9926694 style='height:13.95pt'>&nbsp;</td>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl7426694 colspan=3>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        S&#7889; tài kho&#7843;n / <font class='font1826694'>A/C number:</font>                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=7 class=xl20026694 width=362 style='width:271pt'></td>                                                                                                                                                   	\n");
                htmlStr.Append("                    <td class=xl8826694 width=56 style='width:45.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8826694 width=70 style='width:57.7pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=19 style='width:15.2pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=42 style='width:33.8pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl8726694 width=55 style='width:44.7pt'></td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                                             	\n");
                htmlStr.Append("                    <td height=4 class=xl99266941 style='height:3.0pt'>&nbsp;</td>                                                                                                                                                       	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6726694>&nbsp;</td>                                                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=4>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        H&#7885; tên ng&#432;&#7901;i mua hàng / <font class='font826694'>Buyer</font><font class='font526694'>:</font>                                                                                                  	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=11 class=xl20126694 width=563 style='width:421pt'>&nbsp; " + dt.Rows[0]["buyer"] + "</td>                                                                                                                         	\n");
                htmlStr.Append("                    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=4>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        Tên &#273;&#417;n v&#7883; / <font class='font826694'>Company's name</font><font class='font526694'>                                                                                                             	\n");
                htmlStr.Append("                            :<span style='mso-spacerun:yes'> </span>                                                                                                                                                                     	\n");
                htmlStr.Append("                        </font>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=11 class=xl8826694 width=563 style='width:421pt'> " + dt.Rows[0]["buyerlegalname"] + "</td>                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=3>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        Mã s&#7889; thu&#7871; / <font class='font826694'>Tax code</font><font class='font526694'>                                                                                                                       	\n");
                htmlStr.Append("                            :<span style='mso-spacerun:yes'> </span>                                                                                                                                                                     	\n");
                htmlStr.Append("                        </font>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=12 class=xl21326694> " + dt.Rows[0]["BuyerTaxCode"] + "</td>                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=2>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        &#272;&#7883;a ch&#7881; / <font class='font826694'>Address</font><font class='font526694'>:</font><span style='display:none'>                                                                                   	\n");
                htmlStr.Append("                            <font class='font526694'>                                                                                                                                                                                    	\n");
                htmlStr.Append("                                <span style='mso-spacerun:yes'> </span>                                                                                                                                                                  	\n");
                htmlStr.Append("                            </font>                                                                                                                                                                                                      	\n");
                htmlStr.Append("                        </span>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=13 class=xl20226694 width=652 style='width:488pt'> " + dt.Rows[0]["BuyerAddress"] + "</td>                                                                                                                                   	\n");
                htmlStr.Append("                    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=5>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        &#272;&#7883;a ch&#7881; giao hàng/ <font class='font826694'>Buyer's Address</font><font class='font526694'>:</font><span style='display:none'>                                                                                   	\n");
                htmlStr.Append("                            <font class='font526694'>                                                                                                                                                                                    	\n");
                htmlStr.Append("                                <span style='mso-spacerun:yes'> </span>                                                                                                                                                                  	\n");
                htmlStr.Append("                            </font>                                                                                                                                                                                                      	\n");
                htmlStr.Append("                        </span>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=10 class=xl11626694 width=652 style='width:488pt'> " + dt.Rows[0]["ATTRIBUTE_05"] + "</td>                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");

                htmlStr.Append("                <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=5>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        Hình th&#7913;c thanh toán /<font class='font926694'>                                                                                                                                                            	\n");
                htmlStr.Append("                            Payment method:<span style='mso-spacerun:yes'>                                                                                                                                                               	\n");
                htmlStr.Append("                            </span>                                                                                                                                                                                                      	\n");
                htmlStr.Append("                        </font><font class='font926694'> " + dt.Rows[0]["PaymentMethodCK"] + "</font>                                                                                                                                                    	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=2 class=xl21326694></td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=4>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        S&#7889; tài kho&#7843;n / <font class='font826694'>A/C number</font><font class='font526694'>:</font>                                                                                                           	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=4 class=xl21326694></td>                                                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl6626694 height=20 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=20 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl6626694 colspan=3>                                                                                                                                                                                       	\n");
                htmlStr.Append("                        S&#7889; PO / <font class='font726694'>                                                                                                                                                                          	\n");
                htmlStr.Append("                            PO                                                                                                                                                                                                           	\n");
                htmlStr.Append("                            number                                                                                                                                                                                                       	\n");
                htmlStr.Append("                        </font><font class='font526694'>                                                                                                                                                                                 	\n");
                htmlStr.Append("                            :<span style='mso-spacerun:yes'> </span>                                                                                                                                                                     	\n");
                htmlStr.Append("                        </font>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=12 class=xl21326694>" + dt.Rows[0]["Attribute_04"] + "</td>                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr height=6 style='mso-height-source:userset;height:6.0pt'>                                                                                                                                                            	\n");
                htmlStr.Append("                    <td height=6 class=xl99266941 style='height:6.0pt'>&nbsp;</td>                                                                                                                                                      	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl6526694></td>                                                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl12726694 height=38 style='mso-height-source:userset;height:24.4pt'>                                                                                                                                          	\n");
                htmlStr.Append("                    <td height=38 class=xl12326694 style='height:24.4pt'>&nbsp;</td>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl12426694>STT</td>                                                                                                                                                                                        	\n");
                htmlStr.Append("                    <td colspan=4 class=xl20326694 style='border-right:.5pt solid black; border-left:none'>                                                                                                                              	\n");
                htmlStr.Append("                        Tên hàng hóa, d&#7883;ch v&#7909;                                                                                                                                                                                	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl12426694 style='border-left:none'>&#272;VT</td>                                                                                                                                                          	\n");
                htmlStr.Append("                    <td class=xl14026694 width=56 style='border-left:none;width:45.8pt'>                                                                                                                                                 	\n");
                htmlStr.Append("                        S&#7889;<br>                                                                                                                                                                                                     	\n");
                htmlStr.Append("                        l&#432;&#7907;ng                                                                                                                                                                                                 	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td colspan=2 class=xl20326694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                               	\n");
                htmlStr.Append("                        &#272;&#417;n giá<span style='mso-spacerun:yes'> </span>                                                                                                                                                         	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl12426694 style='border-left:none'>Thành ti&#7873;n</td>                                                                                                                                                  	\n");
                htmlStr.Append("                    <td class=xl12526694 width=56 style='border-left:none;width:45.8pt'>                                                                                                                                                 	\n");
                htmlStr.Append("                        Thu&#7871;                                                                                                                                                                                                       	\n");
                htmlStr.Append("                        su&#7845;t                                                                                                                                                                                                       	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl12526694 width=70 style='width:57.7pt'>Ti&#7873;n thu&#7871;</td>                                                                                                                                        	\n");
                htmlStr.Append("                    <td colspan=3 class=xl20326694 style='border-right:.5pt solid black'>                                                                                                                                                	\n");
                htmlStr.Append("                        T&#7893;ng                                                                                                                                                                                                       	\n");
                htmlStr.Append("                        thành ti&#7873;n                                                                                                                                                                                                 	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl12626694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl13126694 height=37 style='mso-height-source:userset;height:24.3pt'>                                                                                                                                         	\n");
                htmlStr.Append("                    <td height=37 class=xl12826694 style='height:24.3pt'>&nbsp;</td>                                                                                                                                                    	\n");
                htmlStr.Append("                    <td class=xl12926694>No</td>                                                                                                                                                                                         	\n");
                htmlStr.Append("                    <td colspan=4 class=xl14126694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                               	\n");
                htmlStr.Append("                        Description                                                                                                                                                                                                      	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl12926694 style='border-left:none'>Unit</td>                                                                                                                                                              	\n");
                htmlStr.Append("                    <td class=xl12926694 style='border-left:none'>Quantity</td>                                                                                                                                                          	\n");
                htmlStr.Append("                    <td colspan=2 class=xl14126694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                               	\n");
                htmlStr.Append("                        Unit price                                                                                                                                                                                                       	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl12926694 style='border-left:none'>Amount</td>                                                                                                                                                            	\n");
                htmlStr.Append("                    <td class=xl13926694 width=56 style='border-left:none;width:45.8pt'>                                                                                                                                                 	\n");
                htmlStr.Append("                        VAT<br>                                                                                                                                                                                                          	\n");
                htmlStr.Append("                        Rate                                                                                                                                                                                                             	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl13926694 width=70 style='width:57.7pt'>VAT Amount</td>                                                                                                                                                   	\n");
                htmlStr.Append("                    <td colspan=3 class=xl14126694 style='border-right:.5pt solid black'>                                                                                                                                                	\n");
                htmlStr.Append("                        Total                                                                                                                                                                                                            	\n");
                htmlStr.Append("                        Amount                                                                                                                                                                                                           	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl13026694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");
                htmlStr.Append("                <tr class=xl6926694 height=18 style='height:16.7pt'>                                                                                                                                                                     	\n");
                htmlStr.Append("                    <td height=18 class=xl10526694 style='height:16.7pt'>&nbsp;</td>                                                                                                                                                     	\n");
                htmlStr.Append("                    <td class=xl6826694 style='border-top:none'>1</td>                                                                                                                                                                   	\n");
                htmlStr.Append("                    <td colspan=4 class=xl11526694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                               	\n");
                htmlStr.Append("                        2                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl6826694 style='border-top:none;border-left:none'>3</td>                                                                                                                                                  	\n");
                htmlStr.Append("                    <td class=xl6826694 style='border-top:none;border-left:none'>4</td>                                                                                                                                                  	\n");
                htmlStr.Append("                    <td colspan=2 class=xl11526694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                               	\n");
                htmlStr.Append("                        5                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl6826694 style='border-top:none;border-left:none'>6 = 4 x 5</td>                                                                                                                                          	\n");
                htmlStr.Append("                    <td class=xl11526694 style='border-top:none;border-left:none'>7</td>                                                                                                                                                 	\n");
                htmlStr.Append("                    <td class=xl11526694 style='border-top:none'>8</td>                                                                                                                                                                  	\n");
                htmlStr.Append("                    <td colspan=3 class=xl11526694 style='border-right:.5pt solid black'>                                                                                                                                                	\n");
                htmlStr.Append("                        9 = 6 +                                                                                                                                                                                                          	\n");
                htmlStr.Append("                        8                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    </td>                                                                                                                                                                                                                	\n");
                htmlStr.Append("                    <td class=xl10626694>&nbsp;</td>                                                                                                                                                                                     	\n");
                htmlStr.Append("                </tr>                                                                                                                                                                                                                    	\n");

                v_rowHeight = "17.4pt"; //"26.5pt";
                v_rowHeightEmpty = "17.4pt";
                v_rowHeightNumber = 17.4;

                v_rowHeightLast = "17.4pt";// "23.5pt";
                v_rowHeightLastNumber = 17.4;// 23.5;
                v_rowHeightEmptyLast = "17.4pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 32) //!vlongItemName && 
                    {
                        double length_item = dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length / 32;
                        v_rowHeight = "25.5pt"; //"26.5pt";    
                        v_rowHeightLast = "25.5pt"; //"27.5pt";
                        v_rowHeightLastNumber = 25.5;//27.5;
                        v_rowHeightEmptyLast = "17.4pt"; //"23.0pt";
                        v_rowHeightNumber = 25.5;
                        //vlongItemName = true;

                        v_rowHeightNumber = 17.4 + ( Math.Ceiling(length_item) ) * 8.1;
                        v_rowHeightLastNumber = 17.4 + (Math.Ceiling(length_item) ) * 8.1;
                        v_rowHeight = v_rowHeightNumber + "pt"; //"26.5pt";    
                        v_rowHeightLast = v_rowHeightNumber + "pt"; //"27.5pt";

                        //(8.1 * Math.Ceiling(decimal.Parse( dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length) /3)  )
                        //if(Math.Ceiling( dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length  ))
                    }
                    double length_item_m = dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length / 32;
                    if (k == v_countNumberOfPages - 1)
                    {
                        v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                    }
                    else
                    {
                        v_totalHeightPage = v_totalHeightPage - v_rowHeightNumber;
                    }

                    if (dtR == 0)
                    { //dong dau
                    
                        htmlStr.Append("                <tr class=xl7826694 height=24 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                                      	\n");
                        htmlStr.Append("                    <td height=24 class=xl10726694 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                                                                	\n");
                        htmlStr.Append("                    <td class=xl8226694 style='border-top:none'>" + dt_d.Rows[v_index][7] + "&nbsp; </td>                                                                                                                                                             	\n");
                        htmlStr.Append("                    <td colspan=4 height=21 width=232 class=xl14626694 style='border-right:.5pt solid black;                                                                                                                                         	\n");
                        htmlStr.Append("            	  height:17.4pt;width:175pt' align=left valign=top>                                                                                                                                                                                 	\n");
                        htmlStr.Append("                        <![if !vml]><span style='mso-ignore:vglayout;                                                                                                                                                                                	\n");
                        htmlStr.Append("            	  position:absolute;z-index:4;margin-left:52px;margin-top:60px;width:660px;                                                                                                                                                          	\n");
                        htmlStr.Append("            	  height:147.6px'>                                                                                                                                                                                                                     	\n");
                        htmlStr.Append("                            <img width=660 height=147.6                                                                                                                                                                                                	\n");
                        htmlStr.Append("                                 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/CUCKOO_BG_NEW.png'                                                                                                                                            	\n");
                        htmlStr.Append("                                 v:shapes='_x0000_s27865'>                                                                                                                                                                                           	\n");
                        htmlStr.Append("                        </span><![endif]><span style='mso-ignore:vglayout2'>                                                                                                                                                                         	\n");
                        htmlStr.Append("                            <table cellpadding=0 cellspacing=0>                                                                                                                                                                                      	\n");
                        htmlStr.Append("                                <tr>                                                                                                                                                                                                                 	\n");
                        htmlStr.Append("                                    <td colspan=4 height=21 width=232 style='border-right:height:17.4pt;border-left:none;width:175pt'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                      	\n");
                        htmlStr.Append("                                </tr>                                                                                                                                                                                                                	\n");
                        htmlStr.Append("                            </table>                                                                                                                                                                                                                 	\n");
                        htmlStr.Append("                        </span>                                                                                                                                                                                                                      	\n");
                        htmlStr.Append("                    </td>                                                                                                                                                                                                                            	\n");
                        htmlStr.Append("                    <td class=xl7726694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                                                            	\n");
                        htmlStr.Append("                    <td class=xl7726694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                            	\n");
                        htmlStr.Append("                    <td colspan=2 class=xl19526694 style='border-right:.5pt solid black;border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                   	\n");
                        htmlStr.Append("                    <td class=xl7726694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                                            	\n");
                        htmlStr.Append("                    <td class=xl12026694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][13] + "&nbsp;</td>                                                                                                                                          	\n");
                        htmlStr.Append("                    <td class=xl12026694 style='border-top:none'>" + dt_d.Rows[v_index][18] + "&nbsp;</td>                                                                                                                                                           	\n");
                        htmlStr.Append("                    <td colspan=3 class=xl19726694 style='border-right:.5pt solid black'>" + dt_d.Rows[v_index][19] + "&nbsp;</td>                                                                                                                                   	\n");
                        htmlStr.Append("                    <td class=xl10826694 style='border-bottom:none'>&nbsp;</td>                                                                                                                                                                      	\n");
                        htmlStr.Append("                </tr>                                                                                                                                                                                                                                	\n");


                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("                         <tr class=xl7826694 height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                                  	\n");
                            htmlStr.Append("                             <td height=21 class=xl10726694 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                                                            	\n");
                            htmlStr.Append("                             <td class=xl8426694>" + dt_d.Rows[v_index][7] + "&nbsp; </td>                                                                                                                                                                                 	\n");
                            htmlStr.Append("                             <td colspan=4 class=xl15726694 style='border-right:.5pt solid black; border-left:none'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                              	\n");
                            htmlStr.Append("                             <td class=xl8026694 style='border-left:none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                                                                        	\n");
                            htmlStr.Append("                             <td class=xl8026694 style='border-left:none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                                        	\n");
                            htmlStr.Append("                             <td colspan=2 class=xl18326694 style='border-right:.5pt solid black;border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                               	\n");
                            htmlStr.Append("                             <td class=xl8026694 style='border-left:none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                                                        	\n");
                            htmlStr.Append("                             <td class=xl12226694 style='border-left:none'>" + dt_d.Rows[v_index][13] + "&nbsp;</td>                                                                                                                                                      	\n");
                            htmlStr.Append("                             <td class=xl12226694>" + dt_d.Rows[v_index][18] + "&nbsp;</td>																																													 \n");
                            htmlStr.Append("                             <td colspan=3 class=xl18526694 style='border-right:.5pt solid black'>" + dt_d.Rows[v_index][19] + "&nbsp;</td>                                                                                                                                   \n");
                            htmlStr.Append("                             <td class=xl11026694 style='border-bottom:none'>&nbsp;</td>                                                                                                                                                                      \n");
                            htmlStr.Append("                         </tr>                                                                                                                                                                                                                                \n");

                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                               
                                htmlStr.Append("                         <tr class=xl7826694 height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                                  	\n");
                                htmlStr.Append("                             <td height=21 class=xl10726694 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                                                            	\n");
                                htmlStr.Append("                             <td class=xl8426694>" + dt_d.Rows[v_index][7] + "&nbsp; </td>                                                                                                                                                                                 	\n");
                                htmlStr.Append("                             <td colspan=4 class=xl15726694 style='border-right:.5pt solid black; border-left:none'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                              	\n");
                                htmlStr.Append("                             <td class=xl8026694 style='border-left:none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                                                                        	\n");
                                htmlStr.Append("                             <td class=xl8026694 style='border-left:none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                                        	\n");
                                htmlStr.Append("                             <td colspan=2 class=xl18326694 style='border-right:.5pt solid black;border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                               	\n");
                                htmlStr.Append("                             <td class=xl8026694 style='border-left:none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                                                        	\n");
                                htmlStr.Append("                             <td class=xl12226694 style='border-left:none'>" + dt_d.Rows[v_index][13] + "&nbsp;</td>                                                                                                                                                      	\n");
                                htmlStr.Append("                             <td class=xl12226694>" + dt_d.Rows[v_index][18] + "&nbsp;</td>																																													 \n");
                                htmlStr.Append("                             <td colspan=3 class=xl18526694 style='border-right:.5pt solid black'>" + dt_d.Rows[v_index][19] + "&nbsp;</td>                                                                                                                                   \n");
                                htmlStr.Append("                             <td class=xl11026694 style='border-bottom:none'>&nbsp;</td>                                                                                                                                                                      \n");
                                htmlStr.Append("                         </tr>                                                                                                                                                                                                                                \n");

                            }
                            else
                            {
                         
                                htmlStr.Append("                   <tr class=xl7826694 height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                                          \n");
                                htmlStr.Append("                       <td height=21 class=xl10726694 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                                                                    \n");
                                htmlStr.Append("                       <td class=xl8326694 style='border-top:none'>" + dt_d.Rows[v_index][7] + "&nbsp; </td>                                                                                                                                                                 \n");
                                htmlStr.Append("                       <td colspan=4 class=xl14926694 style='border-right:.5pt solid black;border-left:none'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                                       \n");
                                htmlStr.Append("                       <td class=xl7926694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                                                                \n");
                                htmlStr.Append("                       <td class=xl7926694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                                \n");
                                htmlStr.Append("                       <td colspan=2 class=xl15426694 style='border-right:.5pt solid black;border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                       \n");
                                htmlStr.Append("                       <td class=xl7926694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                                                \n");
                                htmlStr.Append("                       <td class=xl12126694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][13] + "&nbsp;</td>                                                                                                                                              \n");
                                htmlStr.Append("                       <td class=xl12126694 style='border-top:none'>" + dt_d.Rows[v_index][18] + "&nbsp;</td>                                                                                                                                                               \n");
                                htmlStr.Append("                       <td colspan=3 class=xl19226694 style='border-right:.5pt solid black'>" + dt_d.Rows[v_index][19] + "&nbsp;</td>                                                                                                                                       \n");
                                htmlStr.Append("                       <td class=xl10926694 style='border-bottom:none'>&nbsp;</td>                                                                                                                                                                          \n");
                                htmlStr.Append("                   </tr>                                                                                                                                                                                                                                    \n");

                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                      
                        htmlStr.Append("                   <tr class=xl7826694 height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                                                          \n");
                        htmlStr.Append("                       <td height=21 class=xl10726694 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                                                                                                    \n");
                        htmlStr.Append("                       <td class=xl8326694 style='border-top:none'>" + dt_d.Rows[v_index][7] + "&nbsp; </td>                                                                                                                                                                 \n");
                        htmlStr.Append("                       <td colspan=4 class=xl14926694 style='border-right:.5pt solid black;border-left:none'>&nbsp;" + dt_d.Rows[v_index][0] + "</td>                                                                                                                       \n");
                        htmlStr.Append("                       <td class=xl7926694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("                       <td class=xl7926694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("                       <td colspan=2 class=xl15426694 style='border-right:.5pt solid black;border-left:none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                       \n");
                        htmlStr.Append("                       <td class=xl7926694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("                       <td class=xl12126694 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][13] + "&nbsp;</td>                                                                                                                                              \n");
                        htmlStr.Append("                       <td class=xl12126694 style='border-top:none'>" + dt_d.Rows[v_index][18] + "&nbsp;</td>                                                                                                                                                               \n");
                        htmlStr.Append("                       <td colspan=3 class=xl19226694 style='border-right:.5pt solid black'>" + dt_d.Rows[v_index][19] + "&nbsp;</td>                                                                                                                                       \n");
                        htmlStr.Append("                       <td class=xl10926694 style='border-bottom:none'>&nbsp;</td>                                                                                                                                                                          \n");
                        htmlStr.Append("                   </tr>                                                                                                                                                                                                                                    \n");

                    }
                    v_index++;
                } //for dtR

                v_spacePerPage = 0;
                if (k < v_countNumberOfPages - 1 && page[k] < rowsPerPage)
                {
                    //for (int i = 0; i < rowsPerPage - page[k]; i++)
                    //{
                    v_spacePerPage += v_totalHeightPage;
                    //}
                }
                else if (k < v_countNumberOfPages - 1 && page[k] == rowsPerPage)
                {
                    v_spacePerPage += v_totalHeightPage;// v_spacePerPage = 10;
                }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("					    <tr class=xl7826694 height=21 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                                                                                                                      \n");
                            htmlStr.Append("					        <td height=21 class=xl10726694 style='height:" + v_rowHeightEmptyLast + "'>&nbsp;</td>                                                                                                                                                                \n");
                            htmlStr.Append("					        <td class=xl8426694>&nbsp;</td>                                                                                                                                                                                                  \n");
                            htmlStr.Append("					        <td colspan=4 class=xl15726694 style='border-right:.5pt solid black;                                                                                                                                                             \n");
                            htmlStr.Append("								  border-left:none'>                                                                                                                                                                                                         \n");
                            htmlStr.Append("					            &nbsp;                                                                                                                                                                                                                       \n");
                            htmlStr.Append("					        </td>                                                                                                                                                                                                                            \n");
                            htmlStr.Append("					        <td class=xl8026694 style='border-left:none'>&nbsp;</td>                                                                                                                                                                         \n");
                            htmlStr.Append("					        <td class=xl8026694 style='border-left:none'>&nbsp;</td>                                                                                                                                                                         \n");
                            htmlStr.Append("					        <td colspan=2 class=xl18326694 style='border-right:.5pt solid black;                                                                                                                                                             \n");
                            htmlStr.Append("								  border-left:none'>                                                                                                                                                                                                         \n");
                            htmlStr.Append("					            &nbsp;                                                                                                                                                                                                                       \n");
                            htmlStr.Append("					        </td>                                                                                                                                                                                                                            \n");
                            htmlStr.Append("					        <td class=xl8026694 style='border-left:none'>&nbsp;</td>                                                                                                                                                                         \n");
                            htmlStr.Append("					        <td class=xl12226694 style='border-left:none'>&nbsp;</td>                                                                                                                                                                        \n");
                            htmlStr.Append("					        <td class=xl12226694>&nbsp;</td>                                                                                                                                                                                                 \n");
                            htmlStr.Append("					        <td colspan=3 class=xl18526694 style='border-right:.5pt solid black'>&nbsp;</td>                                                                                                                                                 \n");
                            htmlStr.Append("					        <td class=xl11026694 style='border-bottom:none'>&nbsp;</td>                                                                                                                                                                      \n");
                            htmlStr.Append("					    </tr>                                                                                                                                                                                                                                \n");

                        }
                        else
                        {
                    
                            htmlStr.Append("			     <tr class=xl7826694 height=21 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                                                                                                                  \n");
                            htmlStr.Append("			         <td height=21 class=xl10726694 style='height:" + v_rowHeightEmptyLast + "'>&nbsp;</td>                                                                                                                                                            \n");
                            htmlStr.Append("			         <td class=xl8326694 style='border-top:none'>&nbsp;</td>                                                                                                                                                                      \n");
                            htmlStr.Append("			         <td colspan=4 class=xl14926694 style='border-right:.5pt solid black;                                                                                                                                                         \n");
                            htmlStr.Append("						  border-left:none'>                                                                                                                                                                                                     \n");
                            htmlStr.Append("			             &nbsp;                                                                                                                                                                                                                   \n");
                            htmlStr.Append("			         </td>                                                                                                                                                                                                                        \n");
                            htmlStr.Append("			         <td class=xl7926694 style='border-top:none;border-left:none'>&nbsp;</td>                                                                                                                                                     \n");
                            htmlStr.Append("			         <td class=xl7926694 style='border-top:none;border-left:none'>&nbsp;</td>                                                                                                                                                     \n");
                            htmlStr.Append("			         <td colspan=2 class=xl15426694 style='border-right:.5pt solid black;                                                                                                                                                         \n");
                            htmlStr.Append("						  border-left:none'>                                                                                                                                                                                                     \n");
                            htmlStr.Append("			             &nbsp;                                                                                                                                                                                                                   \n");
                            htmlStr.Append("			         </td>                                                                                                                                                                                                                        \n");
                            htmlStr.Append("			         <td class=xl7926694 style='border-top:none;border-left:none'>&nbsp;</td>                                                                                                                                                     \n");
                            htmlStr.Append("			         <td class=xl12126694 style='border-top:none;border-left:none'>&nbsp;</td>                                                                                                                                                    \n");
                            htmlStr.Append("			         <td class=xl12126694 style='border-top:none'>&nbsp;</td>                                                                                                                                                                     \n");
                            htmlStr.Append("			         <td colspan=3 class=xl19226694 style='border-right:.5pt solid black'>&nbsp;</td>                                                                                                                                             \n");
                            htmlStr.Append("			         <td class=xl10926694 style='border-bottom:none'>&nbsp;</td>                                                                                                                                                                  \n");
                            htmlStr.Append("			     </tr>                                                                                                                                                                                                                            \n");


                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
             


                    htmlStr.Append("				 <tr class=xl7826694 height=21 style='mso-height-source:userset;height:10pt'>                                                                                                                                                              \n");
                    htmlStr.Append("				     <td height=21 class=xl10726694 style='height:10pt;border-bottom:1.0pt solid #2F75B5;'>&nbsp;</td>                                                                                                                                   \n");
                    htmlStr.Append("				     <td colspan=15 class=xl8326694 style='border-right:none;border-left:none; border-top:.5pt solid windowtext; border-bottom: 1.0pt solid #2F75B5;  '>&nbsp;</td>                                                                                                                                            \n");
                    htmlStr.Append("				     <td class=xl10926694 style='border-bottom:1.0pt solid #2F75B5;'>&nbsp;</td>                                                                                                                                                            \n");
                    htmlStr.Append("				 </tr>                                                                                                                                                                                                                                        \n");
                    htmlStr.Append("            </table>         \n");

                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=18 style='height: " + (v_spacePerPage - 10).ToString() + "pt'>                                                                                                                                                                \n");
                    htmlStr.Append("			<td colspan=27 height=18                                                                                                                                                       \n");
                    htmlStr.Append("				style=' height: " + (v_spacePerPage - 10).ToString() + "pt'>&nbsp;</td>                                                                                                                           \n");
                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>            																																										\n");


                    htmlStr.Append("            <table border=0 cellpadding=0 cellspacing=0 width=795 class=xl6526694																								                                                  	\n");
                    htmlStr.Append("                   style='border-collapse:collapse;table-layout:fixed;width:648.0pt'>                                                                                                                                             	\n");
                    htmlStr.Append("                <col class=xl6526694 width=12 style='mso-width-source:userset;mso-width-alt:426;width:9.8pt'>                                                                                                                     	\n");
                    htmlStr.Append("                <col class=xl6526694 width=38 style='mso-width-source:userset;mso-width-alt:1336;width:30.5pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=81 style='mso-width-source:userset;mso-width-alt:2872;width:66.4pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=48 style='mso-width-source:userset;mso-width-alt:1706;width:39.2pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=41 style='mso-width-source:userset;mso-width-alt:1450;width:33.8pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=62 style='mso-width-source:userset;mso-width-alt:2218;width:51.2pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=49 style='mso-width-source:userset;mso-width-alt:1735;width:40.3pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=56 style='mso-width-source:userset;mso-width-alt:1991;width:45.8pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=34 style='mso-width-source:userset;mso-width-alt:1194;width:27.2pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=42 style='mso-width-source:userset;mso-width-alt:1479;width:33.8pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=78 style='mso-width-source:userset;mso-width-alt:2759;width:63.2pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=56 style='mso-width-source:userset;mso-width-alt:1991;width:45.8pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=70 style='mso-width-source:userset;mso-width-alt:2503;width:57.7pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=19 style='mso-width-source:userset;mso-width-alt:682;width:15.2pt'>                                                                                                                    	\n");
                    htmlStr.Append("                <col class=xl6526694 width=42 style='mso-width-source:userset;mso-width-alt:1479;width:33.8pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=55 style='mso-width-source:userset;mso-width-alt:1962;width:44.7pt'>                                                                                                                   	\n");
                    htmlStr.Append("                <col class=xl6526694 width=12 style='mso-width-source:userset;mso-width-alt:426;width:9.8pt'>                                                                                                                     	\n");


                }


            }// for k                                                                                                                             

            htmlStr.Append("				<tr class=xl13526694 height=34 style='mso-height-source:userset;height:25.95pt'>                                                                                                                                                             \n");
            htmlStr.Append("				    <td height=34 class=xl13226694 style='height:25.95pt'>&nbsp;</td>                                                                                                                                                                        \n");
            htmlStr.Append("				    <td class=xl13326694 style='border-top:none'>&nbsp;</td>                                                                                                                                                                                 \n");
            htmlStr.Append("				    <td class=xl13826694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    <td class=xl13826694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    <td class=xl13826694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    <td class=xl13526694 style='border-top:none'></td>                                                                                                                                                                                       \n");
            htmlStr.Append("				    <td colspan=4 class=xl19126694 width=181 style='width:135pt'>                                                                                                                                                                            \n");
            htmlStr.Append("				        Ti&#7873;n                                                                                                                                                                                                                           \n");
            htmlStr.Append("				        ch&#432;a thu&#7871;<br>                                                                                                                                                                                                             \n");
            htmlStr.Append("				        <font class='font726694'>                                                                                                                                                                                                            \n");
            htmlStr.Append("				            Total amount<span style='mso-spacerun:yes'> </span>                                                                                                                                                                            \n");
            htmlStr.Append("				        </font>                                                                                                                                                                                                                              \n");
            htmlStr.Append("				    </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				    <td colspan=3 class=xl18826694 width=204 style='border-right:.5pt solid black;border-left:none;width:153pt'>                                                                                                                             \n");
            htmlStr.Append("				        Ti&#7873;n thu&#7871; GTGT<br>                                                                                                                                                                                                       \n");
            htmlStr.Append("				        <font class='font726694'>VAT amount<span style='mso-spacerun:yes'> </span></font>                                                                                                                                                    \n");
            htmlStr.Append("				    </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				    <td colspan=3 class=xl18826694 width=116 style='border-right:.5pt solid black;border-left:none;width:86pt'>                                                                                                                              \n");
            htmlStr.Append("				        T&#7893;ng ti&#7873;n <br>                                                                                                                                                                                                           \n");
            htmlStr.Append("				        <font class='font726694'>Total payment</font>                                                                                                                                                                                      \n");
            htmlStr.Append("				    </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				    <td class=xl13426694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				</tr>                                                                                                                                                                                                                                        \n");

            for (int l = 0; l < dt_dvat.Rows.Count; l++)
            {
                htmlStr.Append("				 <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                                         \n");
                htmlStr.Append("				    <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                                    \n");
                htmlStr.Append("				    <td class=xl13626694 colspan=3>&nbsp;" + dt_dvat.Rows[l][1] + "</td>                                                                                                                                                                          \n");
                htmlStr.Append("				    <td class=xl8126694 style='border-top:none'>&nbsp;</td>                                                                                                                                                                              \n");
                htmlStr.Append("				    <td class=xl8126694>&nbsp;</td>                                                                                                                                                                                                      \n");
                htmlStr.Append("				    <td colspan=4 class=xl15626694>" + dt_dvat.Rows[l][3] + "&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				    <td colspan=3 class=xl17726694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                                               \n");
                htmlStr.Append("				        " + dt_dvat.Rows[l][2] + "&nbsp;                                                                                                                                                                                                          \n");
                htmlStr.Append("				    </td>                                                                                                                                                                                                                                \n");
                htmlStr.Append("				    <td colspan=3 class=xl17726694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                                               \n");
                htmlStr.Append("				       " + dt_dvat.Rows[l][4] + "&nbsp;                                                                                                                                                                                                           \n");
                htmlStr.Append("				    </td>                                                                                                                                                                                                                                \n");
                htmlStr.Append("				    <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                                     \n");
                htmlStr.Append("				</tr>                                                                                                                                                                                                                                    \n");

            }
            htmlStr.Append("				    <tr class=xl6626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                                              \n");
            htmlStr.Append("				        <td height=22 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                                        \n");
            htmlStr.Append("				        <td class=xl13726694 colspan=2>&nbsp;T&#7893;ng c&#7897;ng:</td>                                                                                                                                                                         \n");
            htmlStr.Append("				        <td class=xl8126694 style='border-top:none'>&nbsp;</td>                                                                                                                                                                                  \n");
            htmlStr.Append("				        <td class=xl8126694 style='border-top:none'>&nbsp;</td>                                                                                                                                                                                  \n");
            htmlStr.Append("				        <td class=xl8126694 style='border-top:none'>&nbsp;</td>                                                                                                                                                                                  \n");
            htmlStr.Append("				        <td colspan=4 class=xl156266941>" + dt.Rows[0]["netamount_display"] + "&nbsp;</td>                                                                                                                                                                                \n");
            htmlStr.Append("				        <td colspan=3 class=xl18026694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                                                   \n");
            htmlStr.Append("				            " + dt.Rows[0]["vatamount_display"] + "&nbsp;                                                                                                                                                                                                               \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td colspan=3 class=xl18026694 style='border-right:.5pt solid black;border-left:none'>                                                                                                                                                   \n");
            htmlStr.Append("				            " + dt.Rows[0]["totalamount_display"] + "&nbsp;                                                                                                                                                                                                        \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr class=xl6626694 height=21 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                                              \n");
            htmlStr.Append("				        <td height=21 class=xl10226694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                                        \n");
            htmlStr.Append("				        <td class=xl9326694 colspan=5>                                                                                                                                                                                                           \n");
            htmlStr.Append("				            <span style='mso-spacerun:yes'> </span>S&#7889;                                                                                                                                                                                      \n");
            htmlStr.Append("				            ti&#7873;n b&#7857;ng ch&#7919; / <font class='font826694'>Amount in words</font><font class='font526694'> :</font>                                                                                                                  \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td colspan=10 class=xl17326694 width=501 style='border-right:.5pt solid black;width:374pt'>                                                                                                                                             \n");
            htmlStr.Append("				            " + read_prive + "&nbsp;                                                                                                                                                                                                                \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td class=xl10326694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr class=xl7626694 height=22 style='mso-height-source:userset;height:17.4pt'>                                                                                                                                                              \n");
            htmlStr.Append("				        <td height=22 class=xl11126694 style='height:17.4pt'>&nbsp;</td>                                                                                                                                                                        \n");
            htmlStr.Append("				        <td colspan=4 class=xl17526694>Ng&#432;&#7901;i mua hàng / Buyer</td>                                                                                                                                                                    \n");
            htmlStr.Append("				        <td colspan=4 class=xl17526694></td>                                                                                                                                                                                                     \n");
            htmlStr.Append("				        <td colspan=7 class=xl17526694>Ng&#432;&#7901;i bán hàng / Seller</td>                                                                                                                                                                   \n");
            htmlStr.Append("				        <td class=xl11226694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=21 style='height:14.4pt'>                                                                                                                                                                                                         \n");
            htmlStr.Append("				        <td height=21 class=xl9926694 style='height:14.4pt'>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("				        <td colspan=4 class=xl17626694>(Ký, ghi rõ h&#7885; tên)</td>                                                                                                                                                                            \n");
            htmlStr.Append("				        <td colspan=4 class=xl15226694></td>                                                                                                                                                                                                     \n");
            htmlStr.Append("				        <td colspan=7 class=xl15226694>                                                                                                                                                                                                          \n");
            htmlStr.Append("				            Ký, ghi rõ h&#7885; tên<span style='mso-spacerun:yes'> </span>                                                                                                                                                                       \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td class=xl10426694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=21 style='mso-height-source:userset;height:14.4pt'>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td height=21 class=xl9926694 style='height:14.4pt'>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("				        <td colspan=4 class=xl15326694>(Sign &amp; full name)</td>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td colspan=4 class=xl15326694></td>                                                                                                                                                                                                     \n");
            htmlStr.Append("				        <td colspan=7 class=xl15326694>(Signature, full name)</td>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=40 style='mso-height-source:userset;height:10.0pt'>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td height=40 class=xl9926694 style='height:10.0pt'>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=20 style='mso-height-source:userset;height:13.9pt'>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td height=20 class=xl9926694 style='height:13.9pt'>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("																																																																 \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("						<td colspan=7 height=20 class=xl16326694 width=362 style='border-right:.5pt solid black;height:13.9pt;width:270pt' align=left valign=top>                                                                                                \n");
                htmlStr.Append("				            <![if !vml]><span style='mso-ignore:vglayout;position:absolute;z-index:1;margin-left:120px;margin-top:14px;width:46px;height:37px'>                                                                                                  \n");
                htmlStr.Append("				                <img width=46 height=37                                                                                                                                                                                                          \n");
                htmlStr.Append("				                     src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/check_signed.png'                                                                                                                                                        \n");
                htmlStr.Append("				                     v:shapes='Picture_x0020_2'>                                                                                                                                                                                                 \n");
                htmlStr.Append("				            </span><![endif]><span style='mso-ignore:vglayout2'>                                                                                                                                                                                 \n");
                htmlStr.Append("				                <table cellpadding=0 cellspacing=0>                                                                                                                                                                                              \n");
                htmlStr.Append("				                    <tr>                                                                                                                                                                                                                         \n");
                htmlStr.Append("				                        <td colspan=7 height=20 width=362 style='height:13.9pt;width:270pt'>                                                                                                                                                     \n");
                htmlStr.Append("				                            Signature Valid<span style='mso-spacerun:yes'> </span>                                                                                                                                                               \n");
                htmlStr.Append("				                        </td>                                                                                                                                                                                                                    \n");
                htmlStr.Append("				                    </tr>                                                                                                                                                                                                                        \n");
                htmlStr.Append("				                </table>                                                                                                                                                                                                                         \n");
                htmlStr.Append("				            </span>                                                                                                                                                                                                                              \n");
                htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");

            }
            else
            {
                htmlStr.Append("							<td colspan=7 height=20 class=xl16326694 width=362 style='border-right:.5pt solid black;height:13.9pt;width:270pt' align=left valign=top></td>                                                                                       \n");
            }

            htmlStr.Append("				        <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=21 style='height:14.4pt'>                                                                                                                                                                                                         \n");
            htmlStr.Append("				        <td height=21 class=xl9926694 style='height:14.4pt'>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("				        <td class=xl9026694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td colspan=7 class=xl16626694 width=362 style='border-right:.5pt solid black;width:270pt'>                                                                                                                                              \n");
            htmlStr.Append("				            <font class='font1526694'>&#272;&#432;&#7907;c ký b&#7903;i:</font><font class='font1626694'>" + dt.Rows[0]["SignedBy"] + "</font>                                                                                                                \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=20 style='mso-height-source:userset;height:14.4pt'>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td height=20 class=xl9926694 style='height:14.4pt'>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("				        <td class=xl9126694 colspan=2>Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                                                                                                                                                                  \n");
            htmlStr.Append("				        <td colspan=5 class=xl16926694></td>                                                                                                                                                                                                     \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td colspan=7 class=xl17026694 style='border-right:.5pt solid black'>                                                                                                                                                                    \n");
            htmlStr.Append("				            <font class='font1426694'>Ngày ký</font><font class='font1326694'>: " + dt.Rows[0]["SignedDate"] + "</font>                                                                                                                                           \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=21 style='height:14.4pt'>                                                                                                                                                                                                         \n");
            htmlStr.Append("				        <td height=21 class=xl9926694 style='height:14.4pt'>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("				        <td class=xl9226694 colspan=7>                                                                                                                                                                                                           \n");
            htmlStr.Append("				            Tra c&#7913;u t&#7841;i Website: <font class='font1126694'><span style='mso-spacerun:yes'> </span></font><font class='font1226694'>" + dt.Rows[0]["WEBSITE_EI"] + "</font>                                                       \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl9126694 colspan=2>Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "</td>                                                                                                                                                     \n");
            htmlStr.Append("				        <td class=xl6526694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl6526694 colspan=3></td>                                                                                                                                                                                                      \n");
            htmlStr.Append("				        <td class=xl7026694></td>                                                                                                                                                                                                                \n");
            htmlStr.Append("				        <td class=xl10126694>&nbsp;</td>                                                                                                                                                                                                         \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=16 style='mso-height-source:userset;height:12.0pt'>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td height=16 class=xl11326694 style='height:12.0pt'>&nbsp;</td>                                                                                                                                                                         \n");
            htmlStr.Append("				        <td colspan=16 class=xl16026694 style='border-right:1.0pt solid #2F75B5'>                                                                                                                                                                \n");
            htmlStr.Append("				            (C&#7847;n                                                                                                                                                                                                                           \n");
            htmlStr.Append("				            ki&#7875;m tra, &#273;&#7889;i chi&#7871;u khi l&#7853;p, giao, nh&#7853;n                                                                                                                                                           \n");
            htmlStr.Append("				            hóa &#273;&#417;n)                                                                                                                                                                                                                   \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");
            htmlStr.Append("				    <tr height=18 style='mso-height-source:userset;height:12.0pt'>                                                                                                                                                                               \n");
            htmlStr.Append("				        <td height=18 class=xl6526694 style='height:12.0pt'></td>                                                                                                                                                                                \n");
            htmlStr.Append("				        <td colspan=16 class=xl16226694>                                                                                                                                                                                                         \n");
            htmlStr.Append("				            " + dt.Rows[0]["CONTRACT_INFO_EI"] + "                                                                                                                                                                                                                         \n");
            htmlStr.Append("				        </td>                                                                                                                                                                                                                                    \n");
            htmlStr.Append("				    </tr>                                                                                                                                                                                                                                        \n");



            htmlStr.Append("					<![if supportMisalignedColumns]>                                                                                                                        \n");
            htmlStr.Append("					<tr height=0 style='display: none'>                                                                                                                     \n");
            htmlStr.Append("						<td width=12 style='width: 10.35pt'></td>                                                                                                             \n");
            htmlStr.Append("						<td width=38 style='width: 25.2pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=81 style='width: 54.9pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=48 style='width: 41.4pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=41 style='width: 35.65pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=62 style='width: 54.05pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=70 style='width: 59.8pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=70 style='width: 59.8pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=34 style='width: 28.75pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=55 style='width: 47.15pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=102 style='width: 88.55pt'></td>                                                                                                           \n");
            htmlStr.Append("						<td width=19 style='width: 16.1pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=26 style='width: 23pt'></td>                                                                                                              \n");
            htmlStr.Append("						<td width=55 style='width: 47.15pt'></td>                                                                                                            \n");
            htmlStr.Append("						<td width=12 style='width: 10.35pt'></td>                                                                                                             \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<![endif]>                                                                                                                                              \n");
            htmlStr.Append("		</table></td>																															                            \n");
            htmlStr.Append("	</tr>                                                                                                                                                                   \n");
            htmlStr.Append("</table>                                                                                                                                                                    \n");
            htmlStr.Append("</body>                                                                                                                                                                     \n");
            htmlStr.Append("</html>                                                                                                                                                                     \n");

            connection.Close();
            connection.Dispose();

            // using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\\webproject\\e-invoice-ws\\02.Web\\AttachFileText\\" + tei_einvoice_m_pk + ".html"))
            // {
            //     file.WriteLine(htmlStr.ToString()); // "sb" is the StringBuilder
            // }

            return htmlStr.ToString() + "|" + dt.Rows[0]["templateCode"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceSerialNo"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceNumber"];

        }

        public static String NumberToTextVN(decimal total)
        {
            try
            {
                string rs = "";
                if (total.ToString().Substring(0, 1) == "-")
                {
                    rs = "Trừ ";
                }

                total = Math.Round(Math.Abs(total), 0);
                string[] ch = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };
                string[] rch = { "lẻ", "mốt", "", "", "", "lăm" };
                string[] u = { "", "mươi", "trăm", "ngàn", "", "", "triệu", "", "", "tỷ", "", "", "ngàn", "", "", "triệu" };
                string nstr = total.ToString();

                int[] n = new int[nstr.Length];
                int len = n.Length;
                for (int i = 0; i < len; i++)
                {
                    n[len - 1 - i] = Convert.ToInt32(nstr.Substring(i, 1));
                }

                for (int i = len - 1; i >= 0; i--)
                {
                    if (i % 3 == 2)// số 0 ở hàng trăm
                    {
                        if (n[i] == 0 && n[i - 1] == 0 && n[i - 2] == 0) continue;//nếu cả 3 số là 0 thì bỏ qua không đọc
                    }
                    else if (i % 3 == 1) // số ở hàng chục
                    {
                        if (n[i] == 0)
                        {
                            if (n[i - 1] == 0) { continue; }// nếu hàng chục và hàng đơn vị đều là 0 thì bỏ qua.
                            else
                            {
                                rs += " " + rch[n[i]]; continue;// hàng chục là 0 thì bỏ qua, đọc số hàng đơn vị
                            }
                        }
                        if (n[i] == 1)//nếu số hàng chục là 1 thì đọc là mười
                        {
                            rs += " mười"; continue;
                        }
                    }
                    else if (i != len - 1)// số ở hàng đơn vị (không phải là số đầu tiên)
                    {
                        if (n[i] == 0)// số hàng đơn vị là 0 thì chỉ đọc đơn vị
                        {
                            if (i + 2 <= len - 1 && n[i + 2] == 0 && n[i + 1] == 0) continue;
                            rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);
                            continue;
                        }
                        if (n[i] == 1)// nếu là 1 thì tùy vào số hàng chục mà đọc: 0,1: một / còn lại: mốt
                        {
                            rs += " " + ((n[i + 1] == 1 || n[i + 1] == 0) ? ch[n[i]] : rch[n[i]]);
                            rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);
                            continue;
                        }
                        if (n[i] == 5) // cách đọc số 5
                        {
                            if (n[i + 1] != 0) //nếu số hàng chục khác 0 thì đọc số 5 là lăm
                            {
                                rs += " " + rch[n[i]];// đọc số 
                                rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);// đọc đơn vị
                                continue;
                            }
                        }
                    }

                    rs += (rs == "" ? " " : ", ") + ch[n[i]];// đọc số
                    rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);// đọc đơn vị
                }
                if (rs[rs.Length - 1] != ' ')
                    rs += " đồng";
                else
                    rs += "đồng";

                if (rs.Length > 2)
                {
                    string rs1 = rs.Substring(0, 2);
                    rs1 = rs1.ToUpper();
                    rs = rs.Substring(2);
                    rs = rs1 + rs;
                }
                return rs.Trim().Replace("lẻ,", "lẻ").Replace("mươi,", "mươi").Replace("trăm,", "trăm").Replace("mười,", "mười");
            }
            catch
            {
                return "";
            }

        }
        public static string Num2VNText(string s, string ccy)
        {
            //process minus case
            String minus = "";
            if (s.Substring(0, 1) == "-")
            {
                s = s.Replace("-", "").Trim();
                minus = "Trừ ";
            }

            String rtnf = "";
            int l = 0;
            int i = 0;
            int j = 0;
            int dk = 0;
            String[] A = new String[32];
            s = s.Replace(",", "");
            String s1 = "";
            String strTmp = "";
            if (s.Contains("."))
            {
                s1 = s.Substring(s.IndexOf(".") + 1);
                s = s.Substring(0, s.IndexOf("."));
            }
            String[] B = new String[8];
            s = s.Trim();
            l = s.Length;
            //l = s1.length();
            if (l > 32)
            {
                rtnf = "Number Very Large!";
                return rtnf;
            }
            for (i = 0; i < l; i++)
            {
                A[i] = s.Substring(i, 1);
            }
            for (i = 0; i < l; i++)
            {
                if (((l - i) % 3 == 0) && (A[i] == "0") && ((A[i + 1] != "0") || (A[i + 2] != "0")))
                {
                    rtnf += " Không";
                }
                if (A[i] == "2") { rtnf += " Hai"; }
                else
                if (A[i] == "3") { rtnf += " Ba"; }
                else
                if (A[i] == "4") { rtnf += " Bốn"; }
                else
                if (A[i] == "6") { rtnf += " Sáu"; }
                else
                if (A[i] == "7") { rtnf += " Bảy"; }
                else
                if (A[i] == "8") { rtnf += " Tám"; }
                else
                if (A[i] == "9") { rtnf += " Chín"; }
                else
                if (A[i] == "5")
                {
                    if ((i > 0) && ((l - i) % 3 == 1) && (A[i - 1] != "0"))
                    {
                        rtnf += " Lăm";
                    }
                    else
                    {
                        if (i > 0 && (l - i) % 3 == 1 && A[i - 1] != "0")
                        {
                            rtnf += " Lăm";
                        }
                        else
                        {
                            rtnf += " Năm";
                        }
                    }
                }

                if ((i > 2) && (A[i] == "1") && ((l - i) % 3 == 1) && (Int32.Parse(A[i - 1]) > 1))
                {
                    rtnf += " Mốt";
                }
                else if ((A[i] == "1") && (((l - i) % 3) != 2))
                {
                    if ((l - i) % 3 == 1)
                    {
                        if (i > 2 && A[i - 2] == "0" || i < 2 && A[0] == "1" || i > 2 && A[i - 1] == "0" || i > 2 && A[i - 1] == "0")
                        {
                            rtnf += " Một";
                        }
                        else
                        {
                            if (A[i - 1] == "1" || A[i - 1] == "0")
                            {
                                rtnf += " Một";
                            }
                            else
                            {
                                rtnf += " Mốt";
                            }

                        }
                    }
                    else
                    {
                        rtnf += " Một";
                    }
                }


                if (((l - i) % 3) == 2 && A[i] != "0" && A[i] != "1")
                {
                    rtnf += " Mươi";
                }
                else if ((l - i) % 3 == 2 && A[i] != "0")
                {
                    rtnf += " Mười";
                }
                if (i == 0)
                {
                    if ((l - i) % 3 == 2 && A[i] == "0" && A[i + 1] != "0")
                    {
                        rtnf += " Không";
                    }
                }
                else
                {
                    if ((l - i) % 3 == 2 && A[i] == "0" && A[i + 1] != "0")
                    {
                        rtnf += " Lẻ";
                    }
                }
                if ((l - i) % 3 == 0 && (A[i + 1] != "0")) //  || A[i + 2] == "0"
                {
                    rtnf += " Trăm";
                }
                else if ((l - i) % 3 == 0 && A[i] != "0")
                {
                    rtnf += " Trăm";
                }

                if ((l - i) == 4)
                {
                    rtnf += " Nghìn";
                }
                if ((l - i) == 7)
                {
                    rtnf += " Triệu";
                }
                if ((l - i) == 10)
                {
                    rtnf += " Tỷ";
                }
                if ((l - i) == 13)
                {
                    rtnf += " Nghìn Tỷ";
                }
                if ((l - i) == 16)
                {
                    rtnf += " Triệu Tỷ";
                }
                if ((l - i) == 19)
                {
                    rtnf += " Tỷ Tỷ";
                }
                if ((l - i) == 22)
                {
                    rtnf += " Triệu Tỷ Tỷ";
                }
                if ((l - i) == 25)
                {
                    rtnf += " Triệu Tỷ Tỷ";
                }
                if ((l - i) == 28)
                {
                    rtnf += " Tỷ Tỷ Tỷ";
                }
                if ((l - i) % 3 == 0 && A[i] == "0" && A[i + 1] == "0" && A[i + 2] == "0")
                {
                    i = i + 2;
                }
                if ((l - i) % 3 == 1)
                {
                    dk = 1;
                    for (j = i; j < l; j++)
                    {
                        if (A[j] != "0")
                        {
                            dk = 0;
                        }
                    }
                }
                if (dk == 1) break;

            }
            if (ccy == "USD")
            {
                rtnf += " Đô La Mỹ";
                if (s1.Length > 0) //Đọc số lẻ 
                {
                    l = s1.Length;
                    if (l > 8)
                    {
                        rtnf += " ERROR!!!";
                        return rtnf;
                    }
                    for (i = 0; i < l; i++)
                    {
                        B[i] = s1.Substring(i, 1);
                    }
                    strTmp = "";
                    //Dịch Tạm
                    for (i = 0; i < 2; i++)
                    {
                        if ((i > 0) && (B[0] != "0") && (B[0] != "1"))
                        {
                            strTmp += " Mươi";
                        }

                        if (B[i] == "1")
                        {
                            if (i == 0)
                            {
                                strTmp += " Mười";
                            }
                            else
                            {
                                if (B[0] == "1")
                                {
                                    strTmp += " Một";
                                }
                                else
                                {
                                    strTmp += " Mốt";
                                }
                            }
                        }

                        switch (Int32.Parse(B[i]))
                        {

                            case 2:
                                strTmp += " Hai";
                                break;
                            case 3:
                                strTmp += " Ba";
                                break;
                            case 4:
                                strTmp += " Bốn";
                                break;
                            case 5:
                                if (i % 2 == 1 && Int32.Parse(B[0]) > 0)
                                {
                                    strTmp += " Lăm";
                                }
                                else
                                {
                                    strTmp += " Năm";
                                }
                                break;
                            case 6:
                                strTmp += " Sáu";
                                break;
                            case 7:
                                strTmp += " Bảy";
                                break;
                            case 8:
                                strTmp += " Tám";
                                break;
                            case 9:
                                strTmp += " Chín";
                                break;
                        }
                    }
                }
                if (strTmp != "")
                {
                    rtnf = rtnf + " Và" + strTmp + " Cen";
                }
            }

            if (ccy == "VND")
            {
                rtnf += " Đồng.";
            }

            rtnf = minus + rtnf; //process minus case  

            return rtnf;
        }
        public static byte[] ReadWholeArray(Stream stream)
        {
            //Source
            //http://www.yoda.arachsys.com/csharp/readbinary.html
            //Jon Skeet
            byte[] data = new byte[stream.Length];
            int offset = 0;
            int remaining = data.Length;
            while (remaining > 0)
            {
                int read = stream.Read(data, offset, remaining);
                if (read <= 0)
                    throw new EndOfStreamException
                        (String.Format("End of stream reached with {0} bytes left to read", remaining));
                remaining -= read;
                offset += read;
            }

            return data;
        }

        private static string ones(string Number)
        {
            int _Number = Convert.ToInt32(Number);
            String name = "";
            switch (_Number)
            {

                case 1:
                    name = "One";
                    break;
                case 2:
                    name = "Two";
                    break;
                case 3:
                    name = "Three";
                    break;
                case 4:
                    name = "Four";
                    break;
                case 5:
                    name = "Five";
                    break;
                case 6:
                    name = "Six";
                    break;
                case 7:
                    name = "Seven";
                    break;
                case 8:
                    name = "Eight";
                    break;
                case 9:
                    name = "Nine";
                    break;
            }
            return name;
        }

        private static string tens(string Number)
        {
            int _Number = Convert.ToInt32(Number);
            String name = null;
            switch (_Number)
            {
                case 10:
                    name = "Ten";
                    break;
                case 11:
                    name = "Eleven";
                    break;
                case 12:
                    name = "Twelve";
                    break;
                case 13:
                    name = "Thirteen";
                    break;
                case 14:
                    name = "Fourteen";
                    break;
                case 15:
                    name = "Fifteen";
                    break;
                case 16:
                    name = "Sixteen";
                    break;
                case 17:
                    name = "Seventeen";
                    break;
                case 18:
                    name = "Eighteen";
                    break;
                case 19:
                    name = "Nineteen";
                    break;
                case 20:
                    name = "Twenty";
                    break;
                case 30:
                    name = "Thirty";
                    break;
                case 40:
                    name = "Fourty";
                    break;
                case 50:
                    name = "Fifty";
                    break;
                case 60:
                    name = "Sixty";
                    break;
                case 70:
                    name = "Seventy";
                    break;
                case 80:
                    name = "Eighty";
                    break;
                case 90:
                    name = "Ninety";
                    break;
                default:
                    if (_Number > 0)
                    {
                        name = tens(Number.Substring(0, 1) + "0") + " " + ones(Number.Substring(1));
                    }
                    break;
            }
            return name;
        }

        private static string ConvertWholeNumber(string Number)
        {
            string word = "";
            try
            {
                bool beginsZero = false;//tests for 0XX    
                bool isDone = false;//test if already translated    
                int dblAmt = Int32.Parse(Number);//  (Convert.ToDouble(Number));
                                                 //if ((dblAmt > 0) && number.StartsWith("0"))    
                if (dblAmt > 0)
                {//test for zero or digit zero in a nuemric    
                    beginsZero = Number.StartsWith("0");

                    int numDigits = Number.Length;
                    int pos = 0;//store digit grouping    
                    string place = "";//digit grouping name:hundres,thousand,etc...    
                    switch (numDigits)
                    {
                        case 1://ones' range    

                            word = ones(Number);
                            isDone = true;
                            break;
                        case 2://tens' range    
                            word = tens(Number);
                            isDone = true;
                            break;
                        case 3://hundreds' range    
                            pos = (numDigits % 3) + 1;
                            place = " Hundred ";
                            break;
                        case 4://thousands' range    
                        case 5:
                        case 6:
                            pos = (numDigits % 4) + 1;
                            place = " Thousand ";
                            break;
                        case 7://millions' range    
                        case 8:
                        case 9:
                            pos = (numDigits % 7) + 1;
                            place = " Million ";
                            break;
                        case 10://Billions's range    
                        case 11:
                        case 12:

                            pos = (numDigits % 10) + 1;
                            place = " Billion ";
                            break;
                        //add extra case options for anything above Billion...    
                        default:
                            isDone = true;
                            break;
                    }
                    if (!isDone)
                    {//if transalation is not done, continue...(Recursion comes in now!!)    
                        if (Number.Substring(0, pos) != "0" && Number.Substring(pos) != "0")
                        {
                            try
                            {
                                word = ConvertWholeNumber(Number.Substring(0, pos)) + place + ConvertWholeNumber(Number.Substring(pos));
                            }
                            catch { }
                        }
                        else
                        {
                            word = ConvertWholeNumber(Number.Substring(0, pos)) + ConvertWholeNumber(Number.Substring(pos));
                        }

                        //check for trailing zeros    
                        //if (beginsZero) word = " and " + word.Trim();    
                    }
                    //ignore digit grouping names    
                    if (word.Trim().Equals(place.Trim())) word = "";
                }
            }
            catch { }
            return word.Trim();
        }

        private static string ConvertToWords(string numb, string ccy)
        {
            string val = "", wholeNo = numb, points = "", andStr = "", pointStr = "";
            string endStr = "Only", starStr = "US Dollar";
            string amount = "";
            try
            {
                int decimalPlace = numb.IndexOf(".");
                if (decimalPlace > 0)
                {
                    wholeNo = numb.Substring(0, decimalPlace);
                    amount = wholeNo.ToString();
                    points = numb.Substring(decimalPlace + 1);
                    if (Convert.ToInt32(points) > 0)
                    {
                        andStr = "and";// just to separate whole numbers from points/cents    
                        endStr = "Cent";//Cents    
                        pointStr = tens(points);// ConvertDecimals(points);

                        val = string.Format("{0} {1} {2} {3} {4}", starStr, ConvertWholeNumber(amount.Trim()), andStr, endStr, pointStr);// + String.Format("{0}", amount);  //starStr + " " + ConvertWholeNumber(amount).Trim() + " " + andStr + " " + endStr + " " + pointStr;  wholeNo + " " + Convert.ToString(decimalPlace);// ConvertWholeNumber(wholeNo).Trim();//
                    }
                    else
                    {
                        val = String.Format(" {0} {1} ", starStr, ConvertWholeNumber(wholeNo.Trim()));
                    }
                }
                else
                {
                    //wholeNo = numb.Substring(0, decimalPlace);
                    //points = numb.Substring(decimalPlace + 1);
                    //if (Convert.ToInt32(points) > 0)
                    //{
                    if (ccy == "USD")
                    {
                        val = String.Format(" {0} {1} ", starStr, ConvertWholeNumber(wholeNo.Trim()));
                    }
                    else
                    {
                        endStr = "Viet Nam Dong";//Cents    
                        val = String.Format(" {0} {1} ", ConvertWholeNumber(wholeNo.Trim()), endStr);
                    }
                }
            }
            catch
            {

            }
            return val;
        }

        private static string ConvertDecimals(string number)
        {
            string cd = "", digit = "", engOne = "";
            for (int i = 0; i < number.Length; i++)
            {
                digit = number[i].ToString();
                if (digit.Equals("0"))
                {
                    engOne = "Zero";
                }
                else
                {
                    engOne = ones(digit);
                }
                cd += " " + engOne;
            }
            return cd;
        }


    }
}