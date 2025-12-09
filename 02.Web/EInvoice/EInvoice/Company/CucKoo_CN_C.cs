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
    public class CucKoo_CN_C
    {
        public static string View(string tei_einvoice_m_pk, string tei_company_pk, string dbName)
        {
            /*string dbUser = "genuwin", dbPwd = "genuwin2";//NOBLANDBD  EINVOICE_252
            string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);*/
            string _conString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=123.30.104.243)(PORT=1941))(CONNECT_DATA=(SERVER=dedicated)(SERVICE_NAME=NOBLANDBD)));User ID=genuwin;Password=genuwin2";


            string Procedure = "stacfdstac71_r_02_1";
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

            Procedure = "stacfdstac71_r_03";

            command = new OracleCommand(Procedure, connection);
            command.CommandType = CommandType.StoredProcedure;


            command.Parameters.Add("p_tei_einvoice_m_pk", OracleDbType.Varchar2, 1000).Value = tei_einvoice_m_pk;
            command.Parameters.Add("p_user_id", OracleDbType.Varchar2, 1000).Value = "genuwin";
            command.Parameters.Add("p_rtn_value", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
            DataSet ds_d = new DataSet();
            OracleDataAdapter da_d = new OracleDataAdapter(command);
            da_d.Fill(ds_d);
            DataTable dt_d = ds_d.Tables[0];

            int pos = 11, pos_lv = 20, v_count = 0, count_page = 0, count_page_v = 0, r = 0, x = 0;

            v_count = dt_d.Rows.Count;  //_Invoices.Inv[0].Invoice.Products.Product.Count();
            int[] page = new int[10] {0, 0, 0, 0, 0, 0, 0, 0, 0, 0};

            int v_index = -1, rowsPerPage = 20;

            if(v_count % pos_lv == 0)
            {
                int n1 = v_count / pos_lv;
                for(int i = 0; i < n1 - 1; i++)
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
            for(int i = 0; i < page.Length; i++)
            {
                if(page[i] > 0)
                {
                    v_countNumberOfPages++;
                }
            }


            /*
            count_page = v_count / pos;

            if (v_count % pos > 0)
            {
                count_page += 1;
            }

            //	System.out.println(count_page);
            count_page_v = v_count / pos_lv;

            if (v_count % pos_lv > 0)
            {
                count_page_v = count_page_v + 1;
            }

            int count_last = (pos % count_page_v) - v_count;

            if ((v_count % pos_lv > 10) || (v_count % pos_lv == 0 && (v_count % pos_lv) % 2 == 0))
            {
                r++;
            }
            */
            string read_prive = "", read_en = "", read_amount = "", amout_vat = "";
            read_amount = dt.Rows[0]["TotalAmountInWord"].ToString();

            if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            {
                read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmountInWord"].ToString()));
            }
            else
            {
                read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            }
            read_prive = read_prive.Replace(",", "");

            read_prive = read_prive.ToString().Substring(0, 2) + read_prive.ToString().Substring(2, read_prive.Length - 2).ToLower().Replace("mỹ", "Mỹ");

            read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower() + '.';
            
            //read_prive = dt.Rows[0]["amount_word_vie"].ToString();

            //read_en = dt.Rows[0]["amount_word_eng"].ToString();

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
            htmlStr.Append("<!--                                                                                                                                                                        \n");
            htmlStr.Append("table {                                                                                                                                                                     \n");
            htmlStr.Append("	mso-displayed-decimal-separator: '\\.';                                                                                                                                  \n");
            htmlStr.Append("	mso-displayed-thousand-separator: '\\,';                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font529612 {                                                                                                                                                               \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font629612 {                                                                                                                                                               \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font729612 {                                                                                                                                                               \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font829612 {                                                                                                                                                               \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font929612 {                                                                                                                                                               \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1029612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1129612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: #0066CC;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1229612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Times New Roman, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1329612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Times New Roman, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1429612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.35pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1529612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1629612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.35pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1729612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");

            htmlStr.Append(".font1829612 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");

            htmlStr.Append(".font1829613 {                                                                                                                                                              \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif !important;                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6529612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6629612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6729612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6829612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid black;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6929612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7029612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                          \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                          \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                          \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7129612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7229612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Times New Roman, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7329612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7429612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7529612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7629612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7729612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7829612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7929612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8029612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8129612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8229612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8329612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8429612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Times New Roman, sans-serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8529612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8629612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8729612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8829612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8929612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9029612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9129612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9229612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 14.1pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9329612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9429612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9529612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9629612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9729612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9829612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xlborgriditems {                                                                                                                                                                \n");
            htmlStr.Append("	border-top: none !important;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: .7pt solid gray !important;                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9929612 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid #2F75B5;                                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid #2F75B5;                                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid #2F75B5;                                                                                                                                        \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl104296121 {                                                                                                                                                              \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 3.8pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl106296121 {                                                                                                                                                              \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 3.8pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid #2F75B5;                                                                                                                                     \n");
            htmlStr.Append("	border-left: 1.0pt solid #2F75B5;                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid black;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid black;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #FFC000;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 10.35pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #FFC000;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 10.35pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #FFC000;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 10.35pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Times New Roman, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Times New Roman, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Times New Roman, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid #2F75B5;                                                                                                                                     \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid #2F75B5;                                                                                                                                      \n");
            htmlStr.Append("	border-bottom: 1.0pt solid #2F75B5;                                                                                                                                     \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 9.50pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl13929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl14929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");

            htmlStr.Append(".xl15329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");

            htmlStr.Append(".xl15429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                  \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1.0pt solid black;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl16929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 13.5pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17029612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17129612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid black;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1.0pt solid black;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17229612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1.0pt solid black;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17329612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17429612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 19pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid #2F75B5;                                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17529612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 17pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17629612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid #2F75B5;                                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17729612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1.0pt solid #2F75B5;                                                                                                                                        \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17829612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl17929612 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 16.1pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("-->                                                                                                                                                                         \n");
            htmlStr.Append("</style>                                                                                                                                                                    \n");
            htmlStr.Append("</head>                                                                                                                                                                     \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                           \n");
            htmlStr.Append("<table border=0 cellpadding=0 cellspacing=0 width=725 class=xl6529612																			\n");
            htmlStr.Append("		style='border-collapse: collapse; table-layout: fixed; width: 634pt'>                                                                   \n");
            htmlStr.Append("		<tr>                                                                                                                                    \n");
            htmlStr.Append("			<td width='10.2pt'></td>                                                                                                            \n");
            htmlStr.Append("			<td width='624pt'><table border=0 cellpadding=0 cellspacing=0                                                                     \n");
            htmlStr.Append("					width=725 class=xl6529612                                                                                                   \n"); 
            htmlStr.Append("					style='border-collapse: collapse; table-layout: fixed; width: 624pt'>                                                                                 \n");
            htmlStr.Append("					<col class=xl6529612 width=12                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 426; width: 10.35pt'>                                                                                \n");
            htmlStr.Append("					<!-- 1 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=38                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 1336; width: 33.2pt'>                                                                              \n");
            htmlStr.Append("					<!-- 2 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=81                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 2872; width: 66.4pt'>                                                                              \n");
            htmlStr.Append("					<!-- 3 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=48                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 1706; width: 41.4pt'>                                                                              \n");
            htmlStr.Append("					<!-- 4 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=41                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 1450; width: 36.65pt'>                                                                              \n");
            htmlStr.Append("					<!-- 5 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=62                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 2218; width: 66.50pt'>                                                                              \n");
            htmlStr.Append("					<!-- 6 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=70 span=2                                                                                                                    \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 2474; width: 68.8pt'>                                                                              \n");
            htmlStr.Append("					<!-- 7 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=34                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 1194; width: 39.75pt'>                                                                              \n");
            htmlStr.Append("					<!-- 8 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=55                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 1962; width: 47.15pt'>                                                                              \n");
            htmlStr.Append("					<!-- 9 -->                                                                                                                                              \n");
            htmlStr.Append("					<col class=xl6529612 width=102                                                                                                                          \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 3640; width: 80.55pt'>                                                                              \n");
            htmlStr.Append("					<!-- 10 -->                                                                                                                                             \n");
            htmlStr.Append("					<col class=xl6529612 width=19                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 682; width: 16.1pt'>                                                                               \n");
            htmlStr.Append("					<!-- 11 -->                                                                                                                                             \n");
            htmlStr.Append("					<col class=xl6529612 width=26                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 938; width: 23pt'>                                                                                 \n");
            htmlStr.Append("					<!-- 12 -->                                                                                                                                             \n");
            htmlStr.Append("					<col class=xl6529612 width=55                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 1962; width: 47.15pt'>                                                                              \n");
            htmlStr.Append("					<!-- 13 -->                                                                                                                                             \n");
            htmlStr.Append("					<col class=xl6529612 width=12                                                                                                                           \n");
            htmlStr.Append("						style='mso-width-source: userset; mso-width-alt: 426; width: 10.35pt'>                                                                                \n");
            htmlStr.Append("					<!-- 14 -->                                                                                                                                             \n");
            v_index = 0;
            string v_titlePageNumber = "";
            double v_spacePerPage = 0;

            string v_rowHeight = "23.5pt";//"26.5pt";
            string v_rowHeightEmpty = "22.0pt";
            double v_rowHeightNumber = 26.5;

            string v_rowHeightLast = "23.5pt";
            double v_rowHeightLastNumber = 23.5;
            string v_rowHeightEmptyLast = "22.5pt";

            bool vlongItemName = false;

            double v_totalHeightLastPage = 223.5;//  258.5;

            double v_totalHeightPage = 540; //540;

            for ( int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 540;

                if (v_countNumberOfPages > 1)
                {
                    if (k == 0)
                    {
                        v_titlePageNumber = "Trang 1/" + v_countNumberOfPages.ToString();
                    }
                    else if(k < v_countNumberOfPages - 1)
                    {
                        v_titlePageNumber = "tiep theo trang truoc - Trang " + (k + 1).ToString() + "/ " + v_countNumberOfPages.ToString();
                    }
                    else if(k == v_countNumberOfPages - 1)
                    {
                        v_titlePageNumber = "Tiep theo trang truoc - Trang " + (k + 1).ToString() + "/ " + v_countNumberOfPages.ToString();
                    }
                }

                if (k== v_countNumberOfPages - 1)
                {
                    rowsPerPage = pos;
                }
                else
                {
                    rowsPerPage = pos_lv;
                }
                    htmlStr.Append("					<tr height=10 style='mso-height-source: userset; height: 24.77pt'>                                                                                      \n");
                    htmlStr.Append("						<td height=10 class=xl6529612 width=12                                                                                                              \n");
                    htmlStr.Append("							style='height: 24.77pt; width: 10.35pt'></td>                                                                                                       \n");
                    htmlStr.Append("						<td class=xl6529612 width=38 style='width: 32.2pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=81 style='width: 61pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=48 style='width: 41.4pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=41 style='width: 35.65pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=62 style='width: 47pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=70 style='width: 59.8pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=70 style='width: 59.8pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=34 style='width: 28.75pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=55 style='width: 47.15pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=102 style='width: 88.55pt'></td>                                                                                             \n");
                    htmlStr.Append("						<td class=xl6529612 width=19 style='width: 14pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=26 style='width: 23pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=55 style='width: 47.15pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612 width=12 style='width: 10.35pt'></td>                                                                                               \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl7329612 height=22                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 21.1pt'>                                                                                                \n");
                    htmlStr.Append("						<td height=22 class=xl9929612 style='height: 21.1pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl10029612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("						<td class=xl10029612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("						<td class=xl10029612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("						<td colspan=6 rowspan=2 class=xl17429612>HÓA &#272;&#416;N                                                                                          \n");
                    htmlStr.Append("							GIÁ TR&#7882; GIA T&#258;NG</td>                                                                                                                \n");
                    htmlStr.Append("						<td rowspan=2 class=xl17629612>M&#7851;u s&#7889; / <font                                                                                           \n");
                    htmlStr.Append("							class='font729612'>Form</font><font class='font929612'>:</font><font                                                                            \n");
                    htmlStr.Append("							class='font1529612'><span style='mso-spacerun: yes'> </span></font></td>                                                                        \n");
                    htmlStr.Append("						<td colspan=3 rowspan=2 class=xl17729612>" + dt.Rows[0]["templateCode"] + "</td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl10129612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl7329612 height=6                                                                                                                            \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 4.95pt'>                                                                                                 \n");
                    htmlStr.Append("						<td height=6 class=xl10229612 style='height: 4.95pt'>&nbsp;</td>                                                                                    \n");
                    htmlStr.Append("						<td colspan=3 rowspan=3 height=49 class=xl7529612 width=167                                                                                         \n");
                    htmlStr.Append("							style='mso-ignore: colspan-rowspan; height: 37.5pt; width: 125pt'><![if !vml]><span                                                             \n");
                    htmlStr.Append("							style='mso-ignore: vglayout'>                                                                                                                   \n");
                    htmlStr.Append("								<table cellpadding=0 cellspacing=0>                                                                                                         \n");
                    htmlStr.Append("									<tr>                                                                                                                                    \n");
                    htmlStr.Append("										<td width=6 height=2></td>                                                                                                          \n");
                    htmlStr.Append("									</tr>                                                                                                                                   \n");
                    htmlStr.Append("									<tr>                                                                                                                                    \n");
                    htmlStr.Append("										<td></td>                                                                                                                           \n");
                    htmlStr.Append("										<td><img width=155 height=46                                                                                                        \n");
                    htmlStr.Append("											src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/CucKoo_001.png'                                                           \n");
                    htmlStr.Append("											v:shapes='Picture_x0020_6'></td>                                                                                                \n");
                    htmlStr.Append("										<td width=11></td>                                                                                                                  \n");
                    htmlStr.Append("									</tr>                                                                                                                                   \n");
                    htmlStr.Append("									<tr>                                                                                                                                    \n");
                    htmlStr.Append("										<td height=3></td>                                                                                                                  \n");
                    htmlStr.Append("									</tr>                                                                                                                                   \n");
                    htmlStr.Append("								</table>                                                                                                                                    \n");
                    htmlStr.Append("						</span>                                                                                                                                             \n");
                    htmlStr.Append("						<![endif]>                                                                                                                                          \n");
                    htmlStr.Append("							<!--[if !mso & vml]><span style='width:124.8pt;height:37.2pt'></span><![endif]--></td>                                                          \n");
                    htmlStr.Append("						<td class=xl10329612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl7329612 height=22                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 21.1pt'>                                                                                                \n");
                    htmlStr.Append("						<td height=22 class=xl10229612 style='height: 21.1pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td colspan=6 class=xl17929612>VAT INVOICE</td>                                                                                                     \n");
                    htmlStr.Append("						<td rowspan=2 class=xl9529612>Ký hi&#7879;u / <font                                                                                                 \n");
                    htmlStr.Append("							class='font729612'>Serial</font><font class='font929612'>:                                                                                      \n");
                    htmlStr.Append("						</font></td>                                                                                                                                        \n");
                    htmlStr.Append("						<td colspan=3 rowspan=2 class=xl17829612>" + dt.Rows[0]["InvoiceSerialNo"] + "</td>                                                                                          \n");
                    htmlStr.Append("						<td class=xl10329612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl7329612 height=21 style='height: 19.38pt'>                                                                                                  \n");
                    htmlStr.Append("						<td height=21 class=xl10229612 style='height: 19.38pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td colspan=6 class=xl13529612>(HÓA &#272;&#416;N CHUY&#7874;N &#272;&#7892;I T&#7914; HÓA &#272;&#416;N &#272;I&#7878;N T&#7916;)</td>                                                                                                                \n");
                    htmlStr.Append("						<td class=xl10329612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=23 style='height: 20pt'>                                                                                                                   \n");
                    htmlStr.Append("						<td height=23 class=xl10429612 style='height: 20pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl7429612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7429612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7429612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td colspan=6 class=xl13329612>Ngày / <font                                                                                                         \n");
                    htmlStr.Append("							class='font729612'>Date</font><font class='font529612'><span                                                                                    \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> " + dt.Rows[0]["invoiceissueddate_dd"] + "  </span>tháng / </font><font                                                                           \n");
                    htmlStr.Append("							class='font729612'>month</font><font class='font529612'><span                                                                                   \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> " + dt.Rows[0]["invoiceissueddate_mm"] + "  </span>n&#259;m / </font><font                                                                        \n");
                    htmlStr.Append("							class='font729612'>year</font> " + dt.Rows[0]["invoiceissueddate_yyyy"] + "</td>                                                                                                 \n");
                    htmlStr.Append("						<td class=xl9529612>S&#7889; / <font class='font729612'>Invoice                                                                                     \n");
                    htmlStr.Append("								no</font><font class='font929612'>:<span                                                                                                    \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td colspan=3 class=xl16929612 width=100 style='width: 75pt; vertical-align: middle;'>" + dt.Rows[0]["InvoiceNumber"] + " </td>                                                                    \n");
                    htmlStr.Append("						<td class=xl10529612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");

                    htmlStr.Append("					<tr height=17.5 style='height: " + (v_countNumberOfPages > 1 ? "14.5" : "1.0").ToString() + "pt;'>                                                                                                    \n");
                    htmlStr.Append("						<td height=17.5 colspan='14' class=xl104296121                                                                                                         \n");
                    htmlStr.Append("							style='height: " + (v_countNumberOfPages > 1 ? "14.5" : "1.0").ToString() + "pt; text-align: center; font-size: " + (v_countNumberOfPages > 1 ? 12 : 1).ToString() + "pt '>" + v_titlePageNumber + "</td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl106296121>&nbsp;</td>                                                                                                                   \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");

                    htmlStr.Append("					<tr height=13 style='mso-height-source: userset; height: 10.05pt'>                                                                                      \n");
                    htmlStr.Append("						<td height=13 class=xl10429612 style='height: 10.05pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl7429612 style='border-bottom: 1.0pt solid black;'></td>                                                                             \n");
                    htmlStr.Append("						<td class=xl7429612 style='border-bottom: 1.0pt solid black;'></td>                                                                             \n");
                    htmlStr.Append("						<td class=xl7429612 style='border-bottom: 1.0pt solid black;'></td>                                                                             \n");
                    /**/
                    htmlStr.Append("						<td class=xl6529612 style='border-bottom: 1.0pt solid black;'></td>                                                                             \n");
                    htmlStr.Append("						<td class=xl6529612 style='border-bottom: 1.0pt solid black;'></td>                                                                             \n");
                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");

                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl7729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl10529612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=2 style='height: 1.95pt; font-size: 1pt'>                                                                                                    \n");
                    htmlStr.Append("						<td height=2 colspan='14' class=xl104296121                                                                                                         \n");
                    htmlStr.Append("							style='height: 1.95pt; font-size: 1pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl106296121>&nbsp;</td>                                                                                                                   \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=22 style='mso-height-source: userset; height: 21.1pt'>                                                                                      \n");
                    htmlStr.Append("						<td height=22 class=xl10429612 style='height: 21.1pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=3>&#272;&#417;n v&#7883; bán hàng                                                                                       \n");
                    htmlStr.Append("							/ <font class='font729612'>Seller</font><font class='font529612'>:</font>                                                                       \n");
                    htmlStr.Append("						</td>                                                                                                                                               \n");
                    htmlStr.Append("						<td class=xl9229612 colspan=7>" + dt.Rows[0]["Seller_Name"] + "</td>                                                                                                    \n");
                    //htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td align=left valign=top><![if !vml]><span                                                                                                         \n");
                    htmlStr.Append("							style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: -4px; margin-top: 8px; width: 105px; height:96.2px'><img             \n");
                    htmlStr.Append("								width=105 height=96.2                                                                                                                   \n");
                    htmlStr.Append("								src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/CucKoo_002.png'                                                                       \n");
                    htmlStr.Append("								v:shapes='Picture_x0020_3'></span>                                                                                                          \n");
                    htmlStr.Append("						<![endif]><span style='mso-ignore: vglayout2'>                                                                                                      \n");
                    htmlStr.Append("								<table cellpadding=0 cellspacing=0>                                                                                                         \n");
                    htmlStr.Append("									<tr>                                                                                                                                    \n");
                    htmlStr.Append("										<td height=22 class=xl6629612 width=19                                                                                              \n");
                    htmlStr.Append("											style='height: 21.1pt; width: 16.1pt'></td>                                                                                    \n");
                    htmlStr.Append("									</tr>                                                                                                                                   \n");
                    htmlStr.Append("								</table>                                                                                                                                    \n");
                    htmlStr.Append("						</span></td>                                                                                                                                        \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=22 style='mso-height-source: userset; height: 21.1pt'>                                                                                      \n");
                    htmlStr.Append("						<td height=22 class=xl10429612 style='height: 21.1pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=3>Mã s&#7889; thu&#7871; / <font                                                                                        \n");
                    htmlStr.Append("							class='font729612'>Tax code</font><font class='font529612'>:<span                                                                               \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl9029612 colspan=3>" + dt.Rows[0]["Seller_TaxCode"] + "</td>                                                                                               \n");
                    htmlStr.Append("						<td class=xl8729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl8729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl8729612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=22 style='mso-height-source: userset; height: 21.1pt'>                                                                                      \n");
                    htmlStr.Append("						<td height=22 class=xl10429612 style='height: 21.1pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=3>&#272;&#7883;a ch&#7881; / <font                                                                                      \n");
                    htmlStr.Append("							class='font729612'>Address</font><font class='font529612'>:<span                                                                                \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td colspan=7 rowspan=2 class=xl9329612 width=434                                                                                                   \n");
                    htmlStr.Append("							style='width: 325pt'>" + dt.Rows[0]["Seller_Address"] + "</td>                                                                                                         \n");
                    htmlStr.Append("						<td class=xl9129612 width=19 style='width: 16.1pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl9129612 width=26 style='width: 23pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl9129612 width=55 style='width: 47.15pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=21 style='height: 19.38pt'>                                                                                                                  \n");
                    htmlStr.Append("						<td height=21 class=xl10429612 style='height: 19.38pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl9129612 width=19 style='width: 16.1pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl9129612 width=26 style='width: 23pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl9129612 width=55 style='width: 47.15pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=21 style='height: 19.38pt'>                                                                                                                  \n");
                    htmlStr.Append("						<td height=21 class=xl10429612 style='height: 19.38pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl7629612 colspan=2>&#272;i&#7879;n tho&#7841;i / <font                                                                                   \n");
                    htmlStr.Append("							class='font1829612'>Tel:</font><font class='font1729612'><span                                                                                  \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td colspan=3 class=xl9329612 width=173 style='width: 130pt'>" + dt.Rows[0]["Seller_Tel"] + "</td>                                                                    \n");
                    htmlStr.Append("						<td class=xl9329612 width=70 style='width: 59.8pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl9329612 width=34 style='width: 28.75pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl9329612 width=55 style='width: 47.15pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl9329612 width=102 style='width: 88.55pt'></td>                                                                                           \n");
                    htmlStr.Append("						<td class=xl9129612 width=19 style='width: 16.1pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl9129612 width=26 style='width: 23pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl9129612 width=55 style='width: 47.15pt'></td>                                                                                            \n");
                    htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=21 style='height: 19.04ptt'>                                                                                                                   \n");
                    htmlStr.Append("						<td height=21 class=xl10429612 style='height: 19.04ptt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl7629612 colspan=3>S&#7889; tài kho&#7843;n / <font                                                                                      \n");
                    htmlStr.Append("							class='font1829612'>A/C number:</font></td>                                                                                                     \n");
                    htmlStr.Append("						<td colspan=7 class=xl17029612 width=434 style='width: 325pt'></td>                                                                                 \n");
                    htmlStr.Append("						<td class=xl9129612 width=19 style='width: 14pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl9129612 width=26 style='width: 23pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl9129612 width=55 style='width: 47.15pt'></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=2 style='height: 1.95pt; font-size: 1pt'>                                                                                                    \n");
                    htmlStr.Append("						<td height=2 colspan='1' class=xl104296121                                                                                                          \n");
                    htmlStr.Append("							style='height: 1.95pt; font-size: 1pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td height=2 colspan='13' class=xl104296121                                                                                                         \n");
                    htmlStr.Append("							style='height: 1.95pt; border-bottom: 1.0pt solid black; border-left: none'>&nbsp;</td>                                                     \n");
                    htmlStr.Append("						<td class=xl106296121>&nbsp;</td>                                                                                                                   \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl6629612 height=28                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 21.37pt'>                                                                                                 \n");
                    htmlStr.Append("						<td height=28 class=xl10729612 style='height: 21.37pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=4>H&#7885; tên ng&#432;&#7901;i                                                                                         \n");
                    htmlStr.Append("							mua hàng / <font class='font729612'>Buyer</font><font                                                                                           \n");
                    htmlStr.Append("							class='font529612'>:</font>                                                                                                                     \n");
                    htmlStr.Append("						</td>                                                                                                                                               \n");
                    htmlStr.Append("						<td class=xl6629612 colspan='9'>" + dt.Rows[0]["buyer"] + "</td>                                                                                                 \n");
                    htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl6629612 height=28                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 21.37pt'>                                                                                                 \n");
                    htmlStr.Append("						<td height=28 class=xl10729612 style='height: 21.37pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=4>Tên &#273;&#417;n v&#7883; / <font                                                                                    \n");
                    htmlStr.Append("							class='font729612'>Company's name</font><font class='font529612'>:<span                                                                         \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6629612 colspan='9'>" + dt.Rows[0]["buyerlegalname"] + "</td>                                                                                                  \n");
                    htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl6629612 height=28                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 21.37pt'>                                                                                                 \n");
                    htmlStr.Append("						<td height=28 class=xl10729612 style='height: 21.37pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=3>Mã s&#7889; thu&#7871; / <font                                                                                        \n");
                    htmlStr.Append("							class='font729612'>Tax code</font><font class='font529612'>:<span                                                                               \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6629612 colspan='10'>" + dt.Rows[0]["BuyerTaxCode"] + "</td>                                                                                                   \n");
                    htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl6629612 height=28                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 21.37pt'>                                                                                                 \n");
                    htmlStr.Append("						<td height=28 class=xl10729612 style='height: 21.37pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=3>&#272;&#7883;a ch&#7881; / <font                                                                                      \n");
                    htmlStr.Append("							class='font729612'>Address</font><font class='font529612'>:<span                                                                                \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6629612 colspan='10' style='font-size:12.0pt'>" + dt.Rows[0]["BuyerAddress"] + "</td>                                                                            \n");
                    htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");

                    htmlStr.Append("					<tr class=xl6629612 height=21                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 21.37pt'>                                                                                                \n");
                    htmlStr.Append("						<td height=21 class=xl10729612 style='height: 21.37pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=7>Hình th&#7913;c thanh toán /<font                                                                                     \n");
                    htmlStr.Append("							class='font829612'> Payment method:<span                                                                                                        \n");
                    htmlStr.Append("								style='mso-spacerun: yes'>   " + dt.Rows[0]["PaymentMethodCK"] + "  </span></font><font                                                                         \n");
                    htmlStr.Append("							class='font529612'><span style='mso-spacerun: yes'>                   </span></font></td>                                                       \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=3>S&#7889; tài kho&#7843;n / <font                                                                                      \n");
                    htmlStr.Append("							class='font729612'>A/C number</font><font class='font529612'>:</font></td>                                                                      \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6629612></td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");

                    htmlStr.Append("					<tr class=xl6629612 height=28                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: 19.94pt'>                                                                                                 \n");
                    htmlStr.Append("						<td height=28 class=xl10729612 style='height: 19.94pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl6629612 colspan=3>Số PO / <font                                                                                      \n");
                    htmlStr.Append("							class='font729612'>PO number</font><font class='font529612'>:<span                                                                                \n");
                    htmlStr.Append("								style='mso-spacerun: yes'> </span></font></td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl6629612 colspan='10' style='font-size:12.0pt'>" + dt.Rows[0]["Attribute_04"] + "</td>                                                                            \n");
                    htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");

                    htmlStr.Append("					<tr height=2 style='height: 1.95pt; font-size: 1pt'>                                                                                                    \n");
                    htmlStr.Append("						<td height=2 colspan='14' class=xl104296121                                                                                                         \n");
                    htmlStr.Append("							style='height: 1.95pt; font-size: 1pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl106296121>&nbsp;</td>                                                                                                                   \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=21 style='height: 19.38pt'>                                                                                                                  \n");
                    htmlStr.Append("						<td height=21 class=xl10429612 style='height: 19.38pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl6829612>STT</td>                                                                                                                        \n");
                    htmlStr.Append("						<td colspan=5 class=xl6829612 style='border-left: none'>Tên                                                                                         \n");
                    htmlStr.Append("							hàng hóa, d&#7883;ch v&#7909;</td>                                                                                                              \n");
                    htmlStr.Append("						<td class=xl6829612 style='border-left: none'>&#272;VT</td>                                                                                         \n");
                    htmlStr.Append("						<td colspan=2 class=xl17129612                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>S&#7889;                                                                              \n");
                    htmlStr.Append("							l&#432;&#7907;ng</td>                                                                                                                           \n");
                    htmlStr.Append("						<td class=xl6829612 style='border-left: none'>&#272;&#417;n                                                                                         \n");
                    htmlStr.Append("							giá<span style='mso-spacerun: yes'> </span>                                                                                                     \n");
                    htmlStr.Append("						</td>                                                                                                                                               \n");
                    htmlStr.Append("						<td colspan=3 class=xl17129612                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>Thành                                                                                 \n");
                    htmlStr.Append("							ti&#7873;n</td>                                                                                                                                 \n");
                    htmlStr.Append("						<td class=xl10929612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr height=21 style='height: 19.38pt'>                                                                                                                  \n");
                    htmlStr.Append("						<td height=21 class=xl10429612 style='height: 19.38pt'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class=xl6929612>No</td>                                                                                                                         \n");
                    htmlStr.Append("						<td colspan=5 class=xl6929612 style='border-left: none'>Description</td>                                                                            \n");
                    htmlStr.Append("						<td class=xl6929612 style='border-left: none'>Unit</td>                                                                                             \n");
                    htmlStr.Append("						<td colspan=2 class=xl16329612                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>Quantity</td>                                                                         \n");
                    htmlStr.Append("						<td class=xl6929612 style='border-left: none'>Unit price</td>                                                                                       \n");
                    htmlStr.Append("						<td colspan=3 class=xl16329612                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>Amount</td>                                                                           \n");
                    htmlStr.Append("						<td class=xl11029612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("					<tr class=xl7129612 height=18 style='height: 13.8pt'>                                                                                                   \n");
                    htmlStr.Append("						<td height=18 class=xl11129612 style='height: 13.8pt'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class=xl7029612 style='border-top: none'>1</td>                                                                                                 \n");
                    htmlStr.Append("						<td colspan=5 class=xl7029612 style='border-left: none'>2</td>                                                                                      \n");
                    htmlStr.Append("						<td class=xl7029612 style='border-top: none; border-left: none'>3</td>                                                                              \n");
                    htmlStr.Append("						<td colspan=2 class=xl16629612                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>4</td>                                                                                \n");
                    htmlStr.Append("						<td class=xl7029612 style='border-top: none; border-left: none'>5</td>                                                                              \n");
                    htmlStr.Append("						<td colspan=3 class=xl16629612                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>6                                                                                     \n");
                    htmlStr.Append("							= 4 x 5</td>                                                                                                                                    \n");
                    htmlStr.Append("						<td class=xl11229612>&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                           \n");

            for (int dtR = 0; dtR < page[k]; dtR++)
            {
                if(!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                {
                    v_rowHeightLast = "26.5pt"; //"27.5pt";
                    v_rowHeightLastNumber = 27.5;
                    v_rowHeightEmptyLast = "23.0pt";
                    vlongItemName = true;
                }
                if(k== v_countNumberOfPages - 1)
                    {
                        v_totalHeightLastPage = v_totalHeightLastPage - v_rowHeightLastNumber;
                    }
                else
                    {
                        v_totalHeightPage = v_totalHeightPage - v_rowHeightNumber;
                    }
                if (dtR == 0)
                { //dong dau
                    htmlStr.Append("					<tr class=xl8029612 height=24                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: '" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + " >                                                                                                 \n");
                    htmlStr.Append("						<td height=20.7 class=xl11329612 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td class='xl8429612 font1829613 xlborgriditems' style='border-top: none; vertical-align: middle; font-size: 10pt; font-weight:400; '> " + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                                   \n");
                    htmlStr.Append("						<td colspan=5 class='xl15629612 xl8529612 xlborgriditems' height=20.7 width=302                                                                                  \n");
                    htmlStr.Append("							style='width: 227pt' align=left valign=top><![if !vml]><span                                                                   \n");
                    htmlStr.Append("							style='mso-ignore: vglayout; position: absolute; z-index: 4; margin-left: 52px; margin-top: 60px; width: 564px; '><img           \n");
                    htmlStr.Append("								width=564 height=125                                                                                                                      \n");
                    htmlStr.Append("								src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/CucKoo_003.png'                                                                       \n");
                    htmlStr.Append("								v:shapes='_x0000_s14492'></span>                                                                                                            \n");
                    htmlStr.Append("						<![endif]><span style='mso-ignore: vglayout2'>                                                                                                      \n");
                    htmlStr.Append("								<table cellpadding=0 cellspacing=0>                                                                                                         \n");
                    htmlStr.Append("									<tr>                                                                                                                                    \n");
                    htmlStr.Append("										<td colspan=5 class='font1829613' width=302                                                                                                   \n");
                    htmlStr.Append("											style='font-size: 8.0pt; padding: 3px; border-left: none; width: 227pt; vertical-align: middle; text-align:left; '>" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                             \n");
                    htmlStr.Append("									</tr>                                                                                                                                   \n");
                    htmlStr.Append("								</table>                                                                                                                                    \n");
                    htmlStr.Append("						</span></td>                                                                                                                                        \n");
                    htmlStr.Append("						<td class='xl7929612 font1829613 xlborgriditems' style='font-size: 10.0pt; border-left: none; text-align: center; vertical-align: middle; '>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                \n");
                    htmlStr.Append("						<td colspan=2 class='xl15729612 font1829613 xlborgriditems'                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none; vertical-align: middle; font-size: 10pt; font-weight: 400; '>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                  \n");
                    htmlStr.Append("						<td class='xl7929612 font1829613 xlborgriditems' style='font-size:10.0pt; border-left: none; vertical-align: middle; '>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                \n");
                    htmlStr.Append("						<td colspan=3 class='xl15929612 font1829613 xlborgriditems'                                                                                                                     \n");
                    htmlStr.Append("							style='font-size: 10.0pt; border-right: 1.0pt solid black; border-left: none; vertical-align: middle; '>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                  \n");
                    htmlStr.Append("						<td class=xl11429612 style='border-bottom: none; ' >&nbsp;</td>                                                                                                                    \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                }
                else if (dtR == page[k] - 1)//dong cuoi moi trang
                {
                   if(k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("						<tr class=xl8029612 height=21                                                                                                                       \n");
                            htmlStr.Append("						style='mso-height-source: userset; height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                \n");
                            htmlStr.Append("							<td height=21 class=xl11329612 style='height: '" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ">&nbsp;</td>                                                                              \n");
                            htmlStr.Append("							<td class='xl8629612 font1829613' style='border-bottom: 1.0pt solid black; text-align: center; vertical-align: middle; font-size: 10pt; '>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                             \n");
                            htmlStr.Append("							<td colspan=5 class='xl8629612 font1829613' style='border-bottom: 1.0pt solid black; font-size: 8.0pt; padding: 3px; border-left: none; vertical-align: middle; text-align: left; '>" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                              \n");
                            htmlStr.Append("							<td class='xl8229612 font1829613' style='font-size:10.0pt; border-bottom: 1.0pt solid black; border-left: none; text-align: center; vertical-align: middle; '>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                              \n");
                            htmlStr.Append("							<td colspan=2 class='xl15129612 font1829613'                                                                                                                  \n");
                            htmlStr.Append("								style='font-size: 10pt; border-bottom: 1.0pt solid black; border-right: 1.0pt solid black; border-left: none; vertical-align: middle; font-weight: 400; '>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                              \n");
                            htmlStr.Append("							<td class='xl8229612 font1829613' style='border-bottom: 1.0pt solid black; font-size: 10.0pt; border-left: none; text-align: right; vertical-align: middle; '>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                            \n");
                            htmlStr.Append("							<td colspan=3 class='xl15329612 font1829613'                                                                                                                  \n");
                            htmlStr.Append("								style='font-size:10.0pt; border-bottom: 1.0pt solid black; border-right: 1.0pt solid black; border-left: none; text-align: right; vertical-align: middle;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                           \n");
                            htmlStr.Append("							<td class=xl11629612 style='border-top: none; border-bottom: none; '>&nbsp;</td>                                                                                       \n");
                            htmlStr.Append("						</tr>                                                                                                                                               \n");
                        }
                        else // trang cuoi
                        {
                            if(dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("						<tr class=xl8029612 height=21                                                                                                                       \n");
                                htmlStr.Append("						style='mso-height-source: userset; height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                \n");
                                htmlStr.Append("							<td height=21 class=xl11329612 style='height: '" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ">&nbsp;</td>                                                                              \n");
                                htmlStr.Append("							<td class='xl8629612 font1829613' style='font-size: 10pt; border-bottom: 1.0pt solid black; text-align: center; vertical-align: middle; '>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                             \n");
                                htmlStr.Append("							<td colspan=5 class='xl8629612 font1829613' style='border-bottom: 1.0pt solid black; font-size: 8.0pt; padding: 3px; border-left: none; vertical-align: middle; text-align: left; '>" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                              \n");
                                htmlStr.Append("							<td class='xl8229612 font1829613' style='border-bottom: 1.0pt solid black; font-size:10.0pt; border-left: none; text-align: center; vertical-align: middle; '>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                              \n");
                                htmlStr.Append("							<td colspan=2 class='xl15129612 font1829613'                                                                                                                  \n");
                                htmlStr.Append("								style='font-size: 10pt; border-right: 1.0pt solid black; border-left: none; border-bottom: 1.0pt solid black; vertical-align: middle; font-weight: 400; '>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                              \n");
                                htmlStr.Append("							<td class='xl8229612 font1829613' style='font-size: 10.0pt; border-bottom: 1.0pt solid black; border-left: none;text-align: right; vertical-align: middle; '>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("							<td colspan=3 class='xl15329612 font1829613'                                                                                                                  \n");
                                htmlStr.Append("								style='font-size:10.0pt; border-bottom: 1.0pt solid black; border-right: 1.0pt solid black; border-left: none; text-align: right; vertical-align: middle;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                           \n");
                                htmlStr.Append("							<td class=xl11629612 style='border-top: none; border-bottom: none; '>&nbsp;</td>                                                                                       \n");
                                htmlStr.Append("						</tr>                                                                                                                                               \n");

                            }
                            else
                            {
                                htmlStr.Append("						<tr class=xl8029612 height=21                                                                                                                       \n");
                                htmlStr.Append("						style='mso-height-source: userset; height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                \n");
                                htmlStr.Append("							<td height=21 class=xl11329612 style='height: '" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ">&nbsp;</td>                                                                              \n");
                                htmlStr.Append("							<td class='xl8629612 font1829613 xlborgriditems' style='font-size: 10pt; text-align: center; vertical-align: middle; '>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                             \n");
                                htmlStr.Append("							<td colspan=5 class='xl8629612 font1829613 xlborgriditems' style='font-size: 8.0pt; padding: 3px; border-left: none; vertical-align: middle; text-align: left; '>" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                              \n");
                                htmlStr.Append("							<td class='xl8229612 font1829613 xlborgriditems' style='font-size:10.0pt; border-left: none; text-align: center; vertical-align: middle; '>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                              \n");
                                htmlStr.Append("							<td colspan=2 class='xl15129612 font1829613 xlborgriditems'                                                                                                                  \n");
                                htmlStr.Append("								style='font-size: 10pt; border-right: 1.0pt solid black; border-left: none; vertical-align: middle; font-weight: 400; '>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                              \n");
                                htmlStr.Append("							<td class='xl8229612 font1829613 xlborgriditems' style='font-size: 10.0pt; border-left: none; text-align: right; vertical-align: middle; '>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                            \n");
                                htmlStr.Append("							<td colspan=3 class='xl15329612 font1829613 xlborgriditems'                                                                                                                  \n");
                                htmlStr.Append("								style='font-size:10.0pt; border-right: 1.0pt solid black; border-left: none; text-align: right; vertical-align: middle;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                           \n");
                                htmlStr.Append("							<td class=xl11629612 style='border-top: none; border-bottom: none; '>&nbsp;</td>                                                                                       \n");
                                htmlStr.Append("						</tr>                                                                                                                                               \n");

                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
        			                                                                                                                                           
                    htmlStr.Append("					<tr class=xl8029612 height=21                                                                                                                           \n");
                    htmlStr.Append("						style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                \n");
                    htmlStr.Append("						<td height=21 class=xl11329612 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td class='xl8529612 font1829613 xlborgriditems' style='font-size:10.0pt; text-align: center; vertical-align: middle; '>" + dt_d.Rows[v_index][7] + "&nbsp;</td>                                                                                 \n");
                    htmlStr.Append("						<td colspan=5 class='xl8529612 font1829613 xlborgriditems' style='font-size: 8.0pt; padding: 3px; border-left: none; vertical-align: middle; text-align: left; '>" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                  \n");
                    htmlStr.Append("						<td class='xl8129612 font1829613 xlborgriditems' style='font-size: 10.0pt; border-left: none;text-align: center; vertical-align: middle;'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                  \n");
                    htmlStr.Append("						<td colspan=2 class='xl14629612 font1829613 xlborgriditems'                                                                                                                      \n");
                    htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none; vertical-align: middle; font-size: 10pt; font-weight: 400; '>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                  \n");
                    htmlStr.Append("						<td class='xl8129612 font1829613 xlborgriditems' style='font-size: 10.0pt; border-left: none; text-align: right; vertical-align: middle;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                \n");
                    htmlStr.Append("						<td colspan=3 class='xl14829612 font1829613 xlborgriditems'                                                                                                                      \n");
                    htmlStr.Append("							style='font-size:10.0pt; border-right: 1.0pt solid black; border-left: none; text-align: right; vertical-align: middle;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                \n");
                    htmlStr.Append("						<td class=xl11529612 style='border-top: none; border-bottom: none; '>&nbsp;</td>                                                                                           \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("\n");
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
            else if(k < v_countNumberOfPages - 1 && page[k] == rowsPerPage)
            {
                v_spacePerPage = 10;
            }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("							<tr class=xl8029612 height=20.7                                                                                                                   \n");
                            htmlStr.Append("							style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                            \n");
                            htmlStr.Append("								<td height=21 class=xl11329612 style='height: " + v_rowHeightEmptyLast + "'>&nbsp;</td>                                                                          \n");
                            htmlStr.Append("								<td class=xl8629612 style='border-bottom: 1.0pt solid black;' >&nbsp;</td>                                                                                                             \n");
                            htmlStr.Append("								<td colspan=5 class=xl8629612 style='border-left: none; border-bottom: 1.0pt solid black; '>&nbsp;</td>                                                                         \n");
                            htmlStr.Append("								<td class=xl8229612 style='border-left: none; border-bottom: 1.0pt solid black; '>&nbsp;</td>                                                                                   \n");
                            htmlStr.Append("								<td colspan=2 class=xl15129612                                                                                                              \n");
                            htmlStr.Append("									style='border-right: 1.0pt solid black; border-left: none; border-bottom: 1.0pt solid black;'>&nbsp;</td>                                                                   \n");
                            htmlStr.Append("								<td class=xl8229612 style='border-left: none; border-bottom: 1.0pt solid black;'>&nbsp;</td>                                                                                   \n");
                            htmlStr.Append("								<td colspan=3 class=xl15329612                                                                                                              \n");
                            htmlStr.Append("									style='border-right: 1.0pt solid black; border-left: none; border-bottom: 1.0pt solid black;'>&nbsp;</td>                                                                   \n");
                            htmlStr.Append("								<td class=xl11629612 style='border-top: none; border-bottom: none; '>&nbsp;</td>                                                                                   \n");
                            htmlStr.Append("							</tr>                                                                                                                                           \n");
                        }
                        else
                        {
                            htmlStr.Append("							<tr class=xl8029612 height=20.7                                                                                                                   \n");
                            htmlStr.Append("								style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                        \n");
                            htmlStr.Append("								<td height=21 class=xl11329612 style='height: " + v_rowHeightEmptyLast + "'>&nbsp;</td>                                                                          \n");
                            htmlStr.Append("								<td class='xl8529612 xlborgriditems' style=''>&nbsp;</td>                                                                                                             \n");
                            htmlStr.Append("								<td colspan=5 class='xl8529612 xlborgriditems' style='border-left: none; '>&nbsp;</td>                                                                         \n");
                            htmlStr.Append("								<td class='xl8129612 xlborgriditems' style='border-left: none; '>&nbsp;</td>                                                                                   \n");
                            htmlStr.Append("								<td colspan=2 class='xl14629612 xlborgriditems'                                                                                                              \n");
                            htmlStr.Append("									style='border-right: 1.0pt solid black; border-left: none; '>&nbsp;</td>                                                                   \n");
                            htmlStr.Append("								<td class='xl8129612 xlborgriditems' style='border-left: none; '>&nbsp;</td>                                                                                   \n");
                            htmlStr.Append("								<td colspan=3 class='xl14829612 xlborgriditems'                                                                                                              \n");
                            htmlStr.Append("									style='border-right: 1.0pt solid black; border-left: none; '>&nbsp;</td>                                                                   \n");
                            htmlStr.Append("								<td class=xl11529612 style='border-top: none; border-bottom: none; '>&nbsp;</td>                                                                                   \n");
                            htmlStr.Append("							</tr>                                                                                                                                           \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                    if (k < v_countNumberOfPages - 1)
                    {
                        htmlStr.Append("					<tr height=10.05 style='mso-height-source: userset; height: " + (v_spacePerPage).ToString() + "pt'>                                                                                      \n");
                        htmlStr.Append("						<td height=10.05 class=xl10429612 style='height: " + (v_spacePerPage).ToString() + "pt; border-bottom:1.0pt solid #2F75B5;'>&nbsp;</td>                                                                                  \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl6529612 style='border-bottom:1.0pt solid #2F75B5;'></td>                                                                                                                           \n");
                        htmlStr.Append("						<td class=xl10629612 style='border-bottom:1.0pt solid #2F75B5;'>&nbsp;</td>                                                                                                                    \n");
                        htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    }


                }// for k                                                                                                                             
                                                                                                                                            
                       					                                                                                                                                              		                                                                                                                                                 
            htmlStr.Append("					<tr class=xl6629612 height=22                                                                                                                           \n");
            htmlStr.Append("						style='mso-height-source: userset; height: 21.1pt'>                                                                                                \n");
            htmlStr.Append("						<td height=22 class=xl10729612 style='height: 21.1pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td colspan=6 class=xl14429612>&nbsp;</td>                                                                                                          \n");
            htmlStr.Append("						<td colspan=4 class=xl13629612                                                                                                                      \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black'>C&#7897;ng ti&#7873;n                                                                                    \n");
            htmlStr.Append("							hàng / <font class='font729612'>Total Amount</font><font                                                                                        \n");
            htmlStr.Append("							class='font529612'> :<span style='mso-spacerun: yes'>  </span></font>                                                                           \n");
            htmlStr.Append("						</td>                                                                                                                                               \n");
            htmlStr.Append("						<td colspan=3 class=xl13829612                                                                                                                      \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>" + dt.Rows[0]["netamount_display"] + "&nbsp;</td>                                                                 \n");
            htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr class=xl6629612 height=22                                                                                                                           \n");
            htmlStr.Append("						style='mso-height-source: userset; height: 21.1pt'>                                                                                                \n");
            htmlStr.Append("						<td height=22 class=xl10729612 style='height: 21.1pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td colspan=6 class=xl14429612><span                                                                                                                \n");
            htmlStr.Append("							style='mso-spacerun: yes'> </span>Thu&#7871; su&#7845;t GTGT / <font                                                                            \n");
            htmlStr.Append("							class='font729612'>VAT rate</font><font class='font529612'>:<span                                                                               \n");
            htmlStr.Append("								style='mso-spacerun: yes'>          </span></font>" + dt.Rows[0]["TaxRate"] + "&nbsp;</td>                                                                    \n");
            htmlStr.Append("						<td colspan=4 class=xl13629612                                                                                                                      \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black'>Ti&#7873;n thu&#7871;                                                                                    \n");
            htmlStr.Append("							GTGT /<font class='font729612'> VAT Amount</font><font                                                                                          \n");
            htmlStr.Append("							class='font529612'> :<span style='mso-spacerun: yes'> </span></font>                                                                            \n");
            htmlStr.Append("						</td>                                                                                                                                               \n");
            htmlStr.Append("						<td colspan=3 class=xl13829612                                                                                                                      \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>" + dt.Rows[0]["vatamount_display"] + "&nbsp;</td>                                                               \n");
            htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr class=xl6629612 height=22                                                                                                                           \n");
            htmlStr.Append("						style='mso-height-source: userset; height: 21.1pt'>                                                                                                \n");
            htmlStr.Append("						<td height=22 class=xl10729612 style='height: 21.1pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td class=xl9829612 style='border-top: none'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("						<td class=xl8329612 style='border-top: none'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("						<td class=xl8329612 style='border-top: none'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("						<td class=xl8329612 style='border-top: none'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("						<td class=xl8329612 style='border-top: none'>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("						<td colspan=5 class=xl13629612                                                                                                                      \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black'>T&#7893;ng c&#7897;ng                                                                                    \n");
            htmlStr.Append("							ti&#7873;n thanh toán / <font class='font729612'>Total                                                                                          \n");
            htmlStr.Append("								Payment</font><font class='font529612'>:<span                                                                                               \n");
            htmlStr.Append("								style='mso-spacerun: yes'>  </span></font>                                                                                                  \n");
            htmlStr.Append("						</td>                                                                                                                                               \n");
            htmlStr.Append("						<td colspan=3 class=xl13829612                                                                                                                      \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black; border-left: none'>" + dt.Rows[0]["totalamount_display"] + "&nbsp;</td>                                                        \n");
            htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr class=xl6629612 height=33                                                                                                                           \n");
            htmlStr.Append("						style='mso-height-source: userset; height: 22.55pt'>                                                                                                \n");
            htmlStr.Append("						<td height=33 class=xl10729612 style='height: 22.55pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td class=xl9829612 colspan=5><span style='mso-spacerun: yes'> </span>S&#7889;                                                                      \n");
            htmlStr.Append("							ti&#7873;n b&#7857;ng ch&#7919; / <font class='font729612'>Amount                                                                               \n");
            htmlStr.Append("								in words</font><font class='font529612'> :</font></td>                                                                                      \n");
            htmlStr.Append("						<td colspan=8 class=xl14129612 width=431                                                                                                            \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black; width: 322pt'>&nbsp;" + read_prive + "</td>                                                                     \n");
            htmlStr.Append("						<td class=xl10829612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr class=xl7829612 height=22																									\n");
            htmlStr.Append("						style='mso-height-source: userset; height: 15.25pt'>                                                                        \n");
            htmlStr.Append("						<td height=22 class=xl11729612 style='height: 15.25pt'>&nbsp;</td>                                                          \n");
            htmlStr.Append("						<td colspan=4 class=xl14329612>Ng&#432;&#7901;i chuy&#7875;n                                                                \n");
            htmlStr.Append("							&#273;&#7893;i / Converter</td>                                                                                         \n");
            htmlStr.Append("						<td colspan=4 class=xl14329612>Ng&#432;&#7901;i mua hàng /                                                                  \n");
            htmlStr.Append("							Buyer</td>                                                                                                              \n");
            htmlStr.Append("						<td colspan=5 class=xl14329612>Ng&#432;&#7901;i bán hàng /                                                                  \n");
            htmlStr.Append("							Seller</td>                                                                                                             \n");
            htmlStr.Append("						<td class=xl11829612>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("					</tr>                                                                                                                           \n");
            htmlStr.Append("					<tr height=21 style='height: 14.04pt'>                                                                                          \n");
            htmlStr.Append("						<td height=21 class=xl10429612 style='height: 14.04pt'>&nbsp;</td>                                                          \n");
            htmlStr.Append("						<td colspan=4 class=xl13229612>(Ký, ghi rõ h&#7885; tên)</td>                                                               \n");
            htmlStr.Append("						<td colspan=4 class=xl13329612>(Ký, ghi rõ h&#7885; tên)</td>                                                               \n");
            htmlStr.Append("						<td colspan=5 class=xl13329612>Ký, ghi rõ h&#7885; tên<span                                                                 \n");
            htmlStr.Append("							style='mso-spacerun: yes'> </span></td>                                                                                 \n");
            htmlStr.Append("						<td class=xl11029612>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("					</tr>                                                                                                                           \n");
            htmlStr.Append("					<tr height=21 style='height: 14.04pt'>                                                                                          \n");
            htmlStr.Append("						<td height=21 class=xl10429612 style='height: 14.04pt'>&nbsp;</td>                                                          \n");
            htmlStr.Append("						<td colspan=4 class=xl13429612>(Sign &amp; full name)</td>                                                                  \n");
            htmlStr.Append("						<td colspan=4 class=xl13529612>(Sign &amp; full name)</td>                                                                  \n");
            htmlStr.Append("						<td colspan=5 class=xl13529612>(Signature, full name)</td>                                                                  \n");
            htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                            \n");
            htmlStr.Append("					</tr>                                                                                                                           \n"); 
            htmlStr.Append("					<tr height=28 style='mso-height-source: userset; height: 15.05pt'>                                                                                      \n");
            htmlStr.Append("						<td height=28 class=xl10429612 style='height: 15.05pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr height=20 style='mso-height-source: userset; height: 15.45pt'>                                                                                      \n");
            htmlStr.Append("						<td height=20 class=xl10429612 style='height: 15.45pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td class=xl6529612 colspan='4'>" + dt.Rows[0]["convert_name"] + "</td>                                                                                                                           \n");
            //htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            //htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            //htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {                                                                                                                             
           			                                                                                                                                              
                htmlStr.Append("						<td colspan=5 height=20 class=xl12029612 width=257                                                                                                                   \n");
                htmlStr.Append("							style='border-right: 1.0pt solid black; height: 15.45pt; width: 193pt'                                                                           \n");
                htmlStr.Append("							align=left valign=top><![if !vml]><span                                                                                                         \n");
                htmlStr.Append("							style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: 142px; margin-top: 15px; width: 46px; height: 37px'><img              \n");
                htmlStr.Append("								width=46 height=37                                                                                                                          \n");
                htmlStr.Append("								src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/check_signed.png'                                                                        \n");
                htmlStr.Append("								v:shapes='Picture_x0020_2'></span>                                                                                                          \n");
                htmlStr.Append("						<![endif]><span style='mso-ignore: vglayout2'>                                                                                                      \n");
                htmlStr.Append("								<table cellpadding=0 cellspacing=0>                                                                                                         \n");
                htmlStr.Append("									<tr>                                                                                                                                    \n");
                htmlStr.Append("										<td colspan=5 height=20  width=257                                                                                  \n");
                htmlStr.Append("											style='height: 15.45pt; width: 193pt'>Signature Valid<span                                                                      \n");
                htmlStr.Append("											style='mso-spacerun: yes'> </span></td>                                                                                         \n");
                htmlStr.Append("									</tr>                                                                                                                                   \n");
                htmlStr.Append("								</table>                                                                                                                                    \n");
                htmlStr.Append("						</span></td>                                                                                                                                        \n");
            			                                                                                                                                                
            }else {                                                                                                                                     
            			                                                                                                                                                
                htmlStr.Append("							<td colspan=5 height=20 class=xl12029612 width=257                                                                                              \n");
                htmlStr.Append("												style='height: 15.45pt; width: 193pt;border-right:1.0pt solid black;'>Signature Valid<span                              \n");
                htmlStr.Append("												style='mso-spacerun: yes'> </span></td>                                                                                     \n");
           	}                                                                                                                                             
            htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr height=21 style='height: 19.38pt'>                                                                                                                  \n");
            htmlStr.Append("						<td height=21 class=xl10429612 style='height: 19.38pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td class=xl9529612 colspan=2>" + dt.Rows[0]["convert_date"] + "</td>                                                                                                                 \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td colspan=5 class=xl12329612 width=257                                                                                                            \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black; width: 193pt'><font                                                                                      \n");
            htmlStr.Append("							class='font1429612'>&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                                              \n");
            htmlStr.Append("							class='font1629612'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                                                                   \n");
            htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr height=20 style='mso-height-source: userset; height: 15.0pt'>                                                                                       \n");
            htmlStr.Append("						<td height=20 class=xl10429612 style='height: 15.0pt'>&nbsp;</td>                                                                                   \n");
            htmlStr.Append("						<td class=xl9629612 colspan=2>Mã nh&#7853;n hóa                                                                                                     \n");
            htmlStr.Append("							&#273;&#417;n:</td>                                                                                                                             \n");
            htmlStr.Append("						<td class=xl6529612 colspan=3>" + dt.Rows[0]["matracuu"] + "</td>                                                                                                     \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td colspan=5 class=xl12629612                                                                                                                      \n");
            htmlStr.Append("							style='border-right: 1.0pt solid black'><font                                                                                                    \n");
            htmlStr.Append("							class='font1329612'>Ngày ký</font><font class='font1229612'>:                                                                                   \n");
            htmlStr.Append("								" + dt.Rows[0]["SignedDate"] + "</font></td>                                                                                                                      \n");
            htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr height=21 style='height: 19.38pt'>                                                                                                                  \n");
            htmlStr.Append("						<td height=21 class=xl10429612 style='height: 19.38pt'>&nbsp;</td>                                                                                  \n");
            htmlStr.Append("						<td class=xl9729612 colspan=7>Tra c&#7913;u t&#7841;i                                                                                               \n");
            htmlStr.Append("							Website: <font class='font1029612'><span                                                                                                        \n");
            htmlStr.Append("								style='mso-spacerun: yes'> </span></font><font class='font1129612'>" + dt.Rows[0]["WEBSITE_EI"] + "</font>                              \n");
            htmlStr.Append("						</td>                                                                                                                                               \n");
            htmlStr.Append("						<td class=xl6529612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl7229612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl7229612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl7229612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl7229612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl7229612></td>                                                                                                                           \n");
            htmlStr.Append("						<td class=xl10629612>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr height=18 style='mso-height-source: userset; height: 13.5pt'>                                                                                       \n");
            htmlStr.Append("						<td height=18 class=xl11929612 style='height: 13.5pt'>&nbsp;</td>                                                                                   \n");
            htmlStr.Append("						<td colspan=14 class=xl12929612                                                                                                                     \n");
            htmlStr.Append("							style='border-right: 1.0pt solid #2F75B5'>(C&#7847;n                                                                                            \n");
            htmlStr.Append("							ki&#7875;m tra, &#273;&#7889;i chi&#7871;u khi l&#7853;p, giao,                                                                                 \n");
            htmlStr.Append("							nh&#7853;n hóa &#273;&#417;n)</td>                                                                                                              \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
            htmlStr.Append("					<tr height=18 style='mso-height-source: userset; height: 10.25pt'>                                                                                      \n");
            htmlStr.Append("						<td height=18 class=xl6529612 style='height: 10.25pt'></td>                                                                                         \n");
            htmlStr.Append("						<td colspan=14 class=xl13129612>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                        \n");
            htmlStr.Append("					</tr>                                                                                                                                                   \n");
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


            /*using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"E:\\webproject\\E_INVOICE_WS\\02.Web\\AttachFileText\\" + tei_einvoice_m_pk + ".txt"))
            {
                file.WriteLine(htmlStr.ToString()); // "sb" is the StringBuilder
            }*/


            //string filePath = 'C:\\Users\\genuwin\\Desktop\\' + tei_einvoice_m_pk + '.xml';  D:/webproject/e-invoice-ws/02.Web/EInvoice/img
            //string filePath = "D:\\webproject\\e-invoice-ws\\02.Web\\AttachFileXml\\" + tei_einvoice_m_pk + ".xml";
            string filePath = "D:\\webproject\\e-invoice-ws\\02.Web\\AttachFileXml\\" + dt.Rows[0]["templateCode"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceSerialNo"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceNumber"] + ".xml";            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"E:\\webproject\\E_INVOICE_WS\\02.Web\\AttachFileText\\" + tei_einvoice_m_pk + ".txt"))

            /*if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            //if (!File.Exists(filePath))
            //{
            String l_key = "", l_InvoiceType = "", l_templateCode = "", l_InvoiceSerialNo = "", l_InvoiceNumber = "", l_InvoiceName = "", l_InvoiceIssuedDate = "", l_SignedDate = "",
            l_keycompany = "", l_Seller_TaxCode = "", l_Seller_Address2 = "", l_BuyerAddress = "", l_invoiceissueddate_dd = "", l_invoiceissueddate_mm = "", l_invoiceissueddate_yyyy = "",
            l_Seller_Name = "", l_Seller_Fname = "", l_Seller_Address = "", l_Seller_Tel = "", l_Seller_Fax = "", l_Seller_AccountNo = "", l_Seller_AccountNo2 = "", l_Seller_Website = "",
            l_Buyer = "", l_BuyerLegalName = "", l_BuyerTaxCode = "", l_BuyerPhone = "", l_BuyerFax = "", l_BuyerAccountNo = "", seller_taxcode = "",
            l_BuyerPONo = "", l_BuyerCInvoiceNo = "", l_PaymentMethodCK = "", l_item = "", l_LineNumber = "", l_ItemName = "", l_ItemUnit = "", l_ItemQuantity = "",
            l_ItemPrice = "", l_ItemAmount = "", l_ItemTaxRate = "", l_ItemVatAmount = "", l_ItemAttr1 = "", l_ItemAttr2 = "", l_CurrencyCodeUSD = "",
            l_ExchangeRate = "", l_NetAmount = "", l_TaxRate = "", l_VATAmount = "", l_TotalAmount = "", l_TotalAmountInWord = "", l_matracuu = "",
            l_TracuuWebsite = "", l_SignedBy = "", l_SignedInfor = "", l_SignatureValue = "", l_SubjectName = "", l_Certificate = "", l_systemCode = "",
            l_CKyDTu = "", l_Signature = "", l_SignedInfo = "", l_SignatureMethod = "", l_Reference = "", l_Transform = "", l_signature = "", l_CanonicalizationMethod = "",
            l_DigestMethod = "", l_DigestValue = "", l_Modulus = "", l_Exponent = "", l_X509SubjectName = "", l_X509Certificate = "", l_SignatureValue_S = "",
            l_Attribute_01 = "",    l_Attribute_02 = "",        l_Attribute_03 = "",        l_Attribute_04 = "",        l_Attribute_05 = "",            l_Attribute_06 = "",            l_Attribute_07 = "",
            l_Attribute_08 = "",    l_Attribute_09 = "",        l_Attribute_10 = "";
            String xmlStr = "";


                l_key = dt.Rows[0]["key"].ToString();
                l_InvoiceType = dt.Rows[0]["InvoiceType"].ToString();
                l_templateCode = dt.Rows[0]["templateCode"].ToString();
                l_InvoiceSerialNo = dt.Rows[0]["InvoiceSerialNo"].ToString();
                l_InvoiceNumber = dt.Rows[0]["InvoiceNumber"].ToString();
                l_InvoiceName = dt.Rows[0]["InvoiceName"].ToString();
                l_InvoiceIssuedDate = dt.Rows[0]["InvoiceIssuedDate"].ToString();
                l_invoiceissueddate_dd = dt.Rows[0]["invoiceissueddate_dd"].ToString();
                l_invoiceissueddate_mm = dt.Rows[0]["invoiceissueddate_mm"].ToString();
                l_invoiceissueddate_yyyy = dt.Rows[0]["invoiceissueddate_yyyy"].ToString();
                l_SignedDate = dt.Rows[0]["SignedDate"].ToString();
                l_keycompany = dt.Rows[0]["keycompany"].ToString();
                l_Seller_Name = dt.Rows[0]["Seller_Name"].ToString();
                l_Seller_Fname = dt.Rows[0]["Seller_Fname"].ToString();
                l_Seller_TaxCode = dt.Rows[0]["Seller_TaxCode"].ToString();
                l_Seller_Address = dt.Rows[0]["Seller_Address"].ToString();
                l_Seller_Address2 = dt.Rows[0]["Seller_Address2"].ToString();
                l_Seller_Tel = dt.Rows[0]["Seller_Tel"].ToString();
                l_Seller_Fax = dt.Rows[0]["Seller_Fax"].ToString();
                l_Seller_AccountNo = dt.Rows[0]["Seller_AccountNo"].ToString();
                l_Seller_AccountNo2 = dt.Rows[0]["Seller_AccountNo2"].ToString();
                l_Seller_Website = dt.Rows[0]["Seller_Website"].ToString();
                l_Buyer = dt.Rows[0]["Buyer"].ToString();
                l_BuyerLegalName = dt.Rows[0]["BuyerLegalName"].ToString();
                l_BuyerTaxCode = dt.Rows[0]["BuyerTaxCode"].ToString();
                l_BuyerAddress = dt.Rows[0]["BuyerAddress"].ToString();
                l_BuyerPhone = dt.Rows[0]["BuyerPhone"].ToString();
                l_BuyerFax = dt.Rows[0]["BuyerFax"].ToString();
                l_BuyerAccountNo = dt.Rows[0]["BuyerAccountNo"].ToString();
                l_BuyerPONo = dt.Rows[0]["BuyerPONo"].ToString();
                l_BuyerCInvoiceNo = dt.Rows[0]["BuyerCInvoiceNo"].ToString();
                l_PaymentMethodCK = dt.Rows[0]["PaymentMethodCK"].ToString();
                l_CurrencyCodeUSD = dt.Rows[0]["CurrencyCodeUSD"].ToString();
                l_ExchangeRate = dt.Rows[0]["ExchangeRate"].ToString();
                l_NetAmount = dt.Rows[0]["netamount_display"].ToString();
                l_TaxRate = dt.Rows[0]["TaxRate"].ToString();
                l_VATAmount = dt.Rows[0]["vatamount_display"].ToString();
                l_TotalAmount = dt.Rows[0]["totalamount_display"].ToString();

                if (l_CurrencyCodeUSD == "VND")
                {

                    l_TotalAmountInWord = Num2VNText(dt.Rows[0]["totalamountinword"].ToString(), "VND");
                    l_TotalAmountInWord = l_TotalAmountInWord.Substring(0, 2) + l_TotalAmountInWord.Substring(2, l_TotalAmountInWord.Length - 2).ToLower();
                }
                else if (l_CurrencyCodeUSD == "USD")
                {
                    l_TotalAmountInWord = Num2VNText(dt.Rows[0]["totalamountinword"].ToString(), "USD");
                    if (l_keycompany == "241")
                    {
                        l_TotalAmountInWord = l_TotalAmountInWord.Replace("Cen", "Xu").Replace("cen", "Xu");
                    }

                }

                l_TotalAmountInWord = l_TotalAmountInWord.ToString().Substring(0, 2) + l_TotalAmountInWord.ToString().Substring(2, l_TotalAmountInWord.Length - 2).ToLower().Replace("mỹ", "Mỹ");

                l_matracuu = dt.Rows[0]["matracuu"].ToString();
                l_TracuuWebsite = dt.Rows[0]["TracuuWebsite"].ToString();
                l_SignedBy = dt.Rows[0]["SignedBy"].ToString();
                l_SignedInfor = dt.Rows[0]["SignedInfor"].ToString();
                l_SignatureValue = dt.Rows[0]["SignatureValue"].ToString();
                l_SubjectName = "";
                l_Certificate = dt.Rows[0]["Certificate"].ToString();
                l_systemCode = dt.Rows[0]["systemCode"].ToString();
                l_Attribute_01 = dt.Rows[0]["Attribute_01"].ToString();
                l_Attribute_02 = dt.Rows[0]["Attribute_02"].ToString();
                l_Attribute_03 = dt.Rows[0]["Attribute_03"].ToString();
                l_Attribute_04 = dt.Rows[0]["Attribute_04"].ToString();
                l_Attribute_05 = dt.Rows[0]["Attribute_05"].ToString();
                l_Attribute_06 = dt.Rows[0]["Attribute_06"].ToString();
                l_Attribute_07 = dt.Rows[0]["Attribute_07"].ToString();
                l_Attribute_08 = dt.Rows[0]["Attribute_08"].ToString();
                l_Attribute_09 = dt.Rows[0]["Attribute_09"].ToString();
                l_Attribute_10 = dt.Rows[0]["Attribute_10"].ToString();
                l_Signature = dt.Rows[0]["signature"].ToString();                                // cells[78];
                l_SignatureValue_S = dt.Rows[0]["signaturevalue"].ToString();                           // cells[77];
                l_SignatureMethod = dt.Rows[0]["signaturemethod"].ToString();                          // cells[67];
                l_Reference = "#ID1";                                                            // cells[42];
                l_Transform = dt.Rows[0]["transform"].ToString();                                // cells[68];
                l_CanonicalizationMethod = dt.Rows[0]["canonicalizationmethod"].ToString();                   // cells[66];
                l_DigestMethod = dt.Rows[0]["digestmethod"].ToString();                             // cells[69];
                l_DigestValue = dt.Rows[0]["digestvalue"].ToString();                              // cells[70];
                l_SignatureValue = dt.Rows[0]["signaturevalue"].ToString();                           // cells[76];
                l_Modulus = dt.Rows[0]["modulus"].ToString();                                  // cells[71];
                l_Exponent = dt.Rows[0]["exponent"].ToString();                                 // cells[72];
                l_X509SubjectName = dt.Rows[0]["x509subjectname"].ToString();                          // cells[73];
                l_X509Certificate = dt.Rows[0]["x509certificate"].ToString();                          // cells[74];


            XmlDocument doc = new XmlDocument();
            XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.AppendChild(docNode);

            XmlNode rootElement = doc.CreateElement("Invoice");
            doc.AppendChild(rootElement);
            XmlAttribute invoiceAttribute = doc.CreateAttribute("xmlns:xsi");
            invoiceAttribute.Value = "http://www.w3.org/2001/XMLSchema-instance";
            rootElement.Attributes.Append(invoiceAttribute);
            XmlAttribute invoiceAttribute1 = doc.CreateAttribute("xmlns");
            invoiceAttribute1.Value = "http://www.w3.org/2001/XMLSchema-instance";
            rootElement.Attributes.Append(invoiceAttribute1);

            XmlNode content = doc.CreateElement("Content");
            rootElement.AppendChild(content);
            XmlAttribute contentAttribute = doc.CreateAttribute("id");
            contentAttribute.Value = "ID1";
            content.Attributes.Append(contentAttribute);

            //XmlNode Inv = doc.CreateElement("Inv");
            //content.AppendChild(Inv);

            //XmlNode key = doc.CreateElement("key");
            //key.AppendChild(doc.CreateTextNode(l_key));
            //Inv.AppendChild(key);

            //XmlNode content = doc.CreateElement("Invoice1");
            //Inv.AppendChild(Invoice1);


            //InvoiceType elements
            XmlNode InvoiceType = doc.CreateElement("InvoiceType");
            InvoiceType.AppendChild(doc.CreateTextNode(l_InvoiceType));
            content.AppendChild(InvoiceType);

            //templateCode elements
            XmlNode templateCode = doc.CreateElement("templateCode");
            templateCode.AppendChild(doc.CreateTextNode(l_templateCode));
            content.AppendChild(templateCode);

            //InvoiceSerialNo elements
            XmlNode InvoiceSerialNo = doc.CreateElement("InvoiceSerialNo");
            InvoiceSerialNo.AppendChild(doc.CreateTextNode(l_InvoiceSerialNo));
            content.AppendChild(InvoiceSerialNo);

            //InvoiceNumber elements
            XmlNode InvoiceNumber = doc.CreateElement("InvoiceNumber");
            InvoiceNumber.AppendChild(doc.CreateTextNode(l_InvoiceNumber));
            content.AppendChild(InvoiceNumber);

            //InvoiceName elements
            XmlNode InvoiceName = doc.CreateElement("InvoiceName");
            InvoiceName.AppendChild(doc.CreateTextNode(l_InvoiceName));
            content.AppendChild(InvoiceName);

            //InvoiceIssuedDate elements
            XmlNode InvoiceIssuedDate = doc.CreateElement("InvoiceIssuedDate");
            InvoiceIssuedDate.AppendChild(doc.CreateTextNode(l_InvoiceIssuedDate));
            content.AppendChild(InvoiceIssuedDate);

            //InvoiceIssuedDate elements
            XmlNode SignedDate = doc.CreateElement("SignedDate");
            SignedDate.AppendChild(doc.CreateTextNode(l_SignedDate));
            content.AppendChild(SignedDate);

            //keycompany elements
            XmlNode keycompany = doc.CreateElement("keycompany");
            keycompany.AppendChild(doc.CreateTextNode(tei_company_pk));
            content.AppendChild(keycompany);

            //Seller_Name elements
            XmlNode Seller_Name = doc.CreateElement("Seller_Name");
            Seller_Name.AppendChild(doc.CreateTextNode(l_Seller_Name));
            content.AppendChild(Seller_Name);

            //Seller_Fname elements
            XmlNode Seller_Fname = doc.CreateElement("Seller_Fname");
            Seller_Fname.AppendChild(doc.CreateTextNode(l_Seller_Fname));
            content.AppendChild(Seller_Fname);

            //Seller_TaxCode elements
            XmlNode Seller_TaxCode = doc.CreateElement("Seller_TaxCode");
            Seller_TaxCode.AppendChild(doc.CreateTextNode(l_Seller_TaxCode));
            content.AppendChild(Seller_TaxCode);

            //Seller_Address elements
            XmlNode Seller_Address = doc.CreateElement("Seller_Address");
            Seller_Address.AppendChild(doc.CreateTextNode(l_Seller_Address));
            content.AppendChild(Seller_Address);

            //Seller_Address2 elements
            XmlNode Seller_Address2 = doc.CreateElement("Seller_Address2");
            Seller_Address2.AppendChild(doc.CreateTextNode(l_Seller_Address2));
            content.AppendChild(Seller_Address2);

            //Seller_Tel elements
            XmlNode Seller_Tel = doc.CreateElement("Seller_Tel");
            Seller_Tel.AppendChild(doc.CreateTextNode(l_Seller_Tel));
            content.AppendChild(Seller_Tel);

            //Seller_Fax elements
            XmlNode Seller_Fax = doc.CreateElement("Seller_Fax");
            Seller_Fax.AppendChild(doc.CreateTextNode(l_Seller_Fax));
            content.AppendChild(Seller_Fax);

            //Seller_AccountNo elements
            XmlNode Seller_AccountNo = doc.CreateElement("Seller_AccountNo");
            Seller_AccountNo.AppendChild(doc.CreateTextNode(l_Seller_AccountNo));
            content.AppendChild(Seller_AccountNo);

            //Seller_AccountNo2 elements
            XmlNode Seller_AccountNo2 = doc.CreateElement("Seller_AccountNo2");
            Seller_AccountNo2.AppendChild(doc.CreateTextNode(l_Seller_AccountNo2));
            content.AppendChild(Seller_AccountNo2);

            //Seller_Website elements
            XmlNode Seller_Website = doc.CreateElement("Seller_Website");
            Seller_Website.AppendChild(doc.CreateTextNode(l_Seller_Website));
            content.AppendChild(Seller_Website);

            //Buyer elements
            XmlNode Buyer = doc.CreateElement("Buyer");
            Buyer.AppendChild(doc.CreateTextNode(l_Seller_Website));
            content.AppendChild(Buyer);

            //BuyerLegalName elements
            XmlNode BuyerLegalName = doc.CreateElement("BuyerLegalName");
            BuyerLegalName.AppendChild(doc.CreateTextNode(l_BuyerLegalName));
            content.AppendChild(BuyerLegalName);

            //BuyerTaxCode elements
            XmlNode BuyerTaxCode = doc.CreateElement("BuyerTaxCode");
            BuyerTaxCode.AppendChild(doc.CreateTextNode(l_BuyerTaxCode));
            content.AppendChild(BuyerTaxCode);

            //BuyerAddress elements
            XmlNode BuyerAddress = doc.CreateElement("BuyerAddress");
            BuyerAddress.AppendChild(doc.CreateTextNode(l_BuyerAddress));
            content.AppendChild(BuyerAddress);

            //BuyerPhone elements
            XmlNode BuyerPhone = doc.CreateElement("BuyerPhone");
            BuyerPhone.AppendChild(doc.CreateTextNode(l_BuyerPhone));
            content.AppendChild(BuyerPhone);

            //BuyerFax elements
            XmlNode BuyerFax = doc.CreateElement("BuyerFax");
            BuyerFax.AppendChild(doc.CreateTextNode(l_BuyerFax));
            content.AppendChild(BuyerFax);

            //BuyerAccountNo elements
            XmlNode BuyerAccountNo = doc.CreateElement("BuyerAccountNo");
            BuyerAccountNo.AppendChild(doc.CreateTextNode(l_BuyerAccountNo));
            content.AppendChild(BuyerAccountNo);

            //BuyerPONo elements
            XmlNode BuyerPONo = doc.CreateElement("BuyerPONo");
            BuyerPONo.AppendChild(doc.CreateTextNode(l_BuyerPONo));
            content.AppendChild(BuyerPONo);

            //BuyerCInvoiceNo elements
            XmlNode BuyerCInvoiceNo = doc.CreateElement("BuyerCInvoiceNo");
            BuyerCInvoiceNo.AppendChild(doc.CreateTextNode(l_BuyerCInvoiceNo));
            content.AppendChild(BuyerCInvoiceNo);

            //PaymentMethod elements
            XmlNode PaymentMethod = doc.CreateElement("PaymentMethod");
            PaymentMethod.AppendChild(doc.CreateTextNode(l_PaymentMethodCK));
            content.AppendChild(PaymentMethod);

            //ExchangeRate elements
            XmlNode ExchangeRate = doc.CreateElement("ExchangeRate");
            ExchangeRate.AppendChild(doc.CreateTextNode(l_ExchangeRate));
            content.AppendChild(ExchangeRate);

            //NetAmount elements
            XmlNode NetAmount = doc.CreateElement("NetAmount");
            NetAmount.AppendChild(doc.CreateTextNode(l_NetAmount));
            content.AppendChild(NetAmount);

            //TaxRate elements
            XmlNode TaxRate = doc.CreateElement("TaxRate");
            TaxRate.AppendChild(doc.CreateTextNode(l_TaxRate));
            content.AppendChild(TaxRate);

            //VATAmount elements
            XmlNode VATAmount = doc.CreateElement("VATAmount");
            VATAmount.AppendChild(doc.CreateTextNode(l_VATAmount));
            content.AppendChild(VATAmount);

            //TotalAmountInWord elements
            XmlNode TotalAmountInWord = doc.CreateElement("TotalAmountInWord");
            TotalAmountInWord.AppendChild(doc.CreateTextNode(l_TotalAmountInWord));
            content.AppendChild(TotalAmountInWord);

            //TotalAmount elements
            XmlNode TotalAmount = doc.CreateElement("TotalAmount");
            TotalAmount.AppendChild(doc.CreateTextNode(l_TotalAmount));
            content.AppendChild(TotalAmount);

            //matracuu elements
            XmlNode matracuu = doc.CreateElement("matracuu");
            matracuu.AppendChild(doc.CreateTextNode(l_matracuu));
            content.AppendChild(matracuu);

            //TracuuWebsite elements
            XmlNode TracuuWebsite = doc.CreateElement("TracuuWebsite");
            TracuuWebsite.AppendChild(doc.CreateTextNode(l_TracuuWebsite));
            content.AppendChild(TracuuWebsite);

            //SignedBy elements
            XmlNode SignedBy = doc.CreateElement("SignedBy");
            SignedBy.AppendChild(doc.CreateTextNode(l_SignedBy));
            content.AppendChild(SignedBy);

            //SignedInfor elements
            XmlNode SignedInfor = doc.CreateElement("SignedInfor");
            SignedInfor.AppendChild(doc.CreateTextNode(l_SignedInfor));
            content.AppendChild(SignedInfor);

            //SignatureValue elements
            //XmlNode SignatureValue = doc.CreateElement("SignatureValue");
            //SignatureValue.AppendChild(doc.CreateTextNode(l_SignatureValue));
            //content.AppendChild(SignatureValue);

            //SubjectName elements
            XmlNode SubjectName = doc.CreateElement("SubjectName");
            SubjectName.AppendChild(doc.CreateTextNode(l_SubjectName));
            content.AppendChild(SubjectName);

            //Certificate elements
            XmlNode Certificate = doc.CreateElement("Certificate");
            Certificate.AppendChild(doc.CreateTextNode(l_Certificate));
            content.AppendChild(Certificate);

            //systemCode elements
            XmlNode systemCode = doc.CreateElement("systemCode");
            systemCode.AppendChild(doc.CreateTextNode(l_systemCode));
            content.AppendChild(systemCode);

            //Products elements
            XmlNode Products = doc.CreateElement("Products");
            content.AppendChild(Products);

            for (int dtR = 0; dtR < v_count; dtR++)
            {
                //Product elements
                XmlNode Product = doc.CreateElement("Product");
                Products.AppendChild(Product);

                //LineNumber elements
                XmlNode LineNumber = doc.CreateElement("LineNumber");
                LineNumber.AppendChild(doc.CreateTextNode(dt_d.Rows[dtR][7].ToString()));
                Product.AppendChild(LineNumber);

                //ItemName elements
                XmlNode ItemName = doc.CreateElement("ItemName");
                ItemName.AppendChild(doc.CreateTextNode(dt_d.Rows[dtR][0].ToString()));
                Product.AppendChild(ItemName);

                //ItemUnit elements
                XmlNode ItemUnit = doc.CreateElement("ItemUnit");
                ItemUnit.AppendChild(doc.CreateTextNode(dt_d.Rows[dtR][1].ToString()));
                Product.AppendChild(ItemUnit);

                //ItemQuantity elements
                XmlNode ItemQuantity = doc.CreateElement("ItemQuantity");
                ItemQuantity.AppendChild(doc.CreateTextNode(dt_d.Rows[dtR][2].ToString()));
                Product.AppendChild(ItemQuantity);

                //ItemPrice elements
                XmlNode ItemPrice = doc.CreateElement("ItemPrice");
                ItemPrice.AppendChild(doc.CreateTextNode(dt_d.Rows[dtR][3].ToString()));
                Product.AppendChild(ItemPrice);

                //ItemAmount elements
                XmlNode ItemAmount = doc.CreateElement("ItemAmount");
                ItemAmount.AppendChild(doc.CreateTextNode(dt_d.Rows[dtR][4].ToString()));
                Product.AppendChild(ItemAmount);
            }
            //==================================================================================================================================================  
            // phần này thêm phần chứng thực vào, mới đầu chỉ có Vina và Solution sử dụng 
            // CKyDTu elements
            XmlNode CKyDTu = doc.CreateElement("CKyDTu");
            rootElement.AppendChild(CKyDTu);

            // Signature elements
            XmlNode Signature = doc.CreateElement("Signature");
            CKyDTu.AppendChild(Signature);
            XmlAttribute att_Signature = doc.CreateAttribute("xmlns");
            att_Signature.Value = l_Signature;
            Signature.Attributes.Append(att_Signature);


            // SignedInfo elements
            XmlNode SignedInfo = doc.CreateElement("SignedInfo");
            Signature.AppendChild(SignedInfo);

            // CanonicalizationMethod elements
            XmlNode CanonicalizationMethod = doc.CreateElement("CanonicalizationMethod");
            XmlAttribute att_CanonicalizationMethod = doc.CreateAttribute("Algorithm");
            att_CanonicalizationMethod.Value = l_CanonicalizationMethod;
            CanonicalizationMethod.Attributes.Append(att_CanonicalizationMethod);
            SignedInfo.AppendChild(CanonicalizationMethod);

            // SignatureMethod elements
            XmlNode SignatureMethod = doc.CreateElement("SignatureMethod");
            XmlAttribute att_SignatureMethod = doc.CreateAttribute("Algorithm");
            att_SignatureMethod.Value = l_SignatureMethod;
            SignatureMethod.Attributes.Append(att_SignatureMethod);
            SignedInfo.AppendChild(SignatureMethod);

            // Reference elements
            XmlNode Reference = doc.CreateElement("Reference");
            XmlAttribute att_Reference = doc.CreateAttribute("Algorithm");
            att_Reference.Value = l_Reference;
            Reference.Attributes.Append(att_Reference);
            SignedInfo.AppendChild(Reference);

            // Transforms elements
            XmlNode Transforms = doc.CreateElement("Transforms");
            Reference.AppendChild(Transforms);

            // Transform elements
            XmlNode Transform = doc.CreateElement("Transform");
            XmlAttribute att_Transform = doc.CreateAttribute("Algorithm");
            att_Transform.Value = l_Transform;
            Transform.Attributes.Append(att_Transform);
            Transforms.AppendChild(Transform);

            // DigestMethod elements
            XmlNode DigestMethod = doc.CreateElement("DigestMethod");
            XmlAttribute att_DigestMethod = doc.CreateAttribute("Algorithm");
            att_DigestMethod.Value = l_DigestMethod;
            DigestMethod.Attributes.Append(att_DigestMethod);
            Reference.AppendChild(DigestMethod);

            // DigestValue elements
            XmlNode DigestValue = doc.CreateElement("DigestValue");
            DigestValue.AppendChild(doc.CreateTextNode(l_DigestValue));
            Reference.AppendChild(DigestValue);


            // SignatureValue elements
            XmlNode SignatureValue = doc.CreateElement("SignatureValue");
            SignatureValue.AppendChild(doc.CreateTextNode(l_SignatureValue_S));
            Signature.AppendChild(SignatureValue);

            // KeyInfo elements
            XmlNode KeyInfo = doc.CreateElement("KeyInfo");
            Signature.AppendChild(KeyInfo);


            // KeyValue elements
            XmlNode KeyValue = doc.CreateElement("KeyValue");
            KeyInfo.AppendChild(KeyValue);

            // RSAKeyValue elements
            XmlNode RSAKeyValue = doc.CreateElement("RSAKeyValue");
            KeyValue.AppendChild(RSAKeyValue);

            // Modulus elements
            XmlNode Modulus = doc.CreateElement("Modulus");
            Modulus.AppendChild(doc.CreateTextNode(l_Modulus));
            RSAKeyValue.AppendChild(Modulus);

            // Exponent elements
            XmlNode Exponent = doc.CreateElement("Exponent");
            Exponent.AppendChild(doc.CreateTextNode(l_Exponent));
            RSAKeyValue.AppendChild(Exponent);

            // X509Data elements
            XmlNode X509Data = doc.CreateElement("X509Data");
            KeyInfo.AppendChild(X509Data);


            // X509SubjectName elements
            XmlNode X509SubjectName = doc.CreateElement("X509SubjectName");
            X509SubjectName.AppendChild(doc.CreateTextNode(l_X509SubjectName));
            X509Data.AppendChild(X509SubjectName);

            // X509Certificate elements
            XmlNode X509Certificate = doc.CreateElement("X509Certificate");
            X509Certificate.AppendChild(doc.CreateTextNode(l_X509Certificate));
            X509Data.AppendChild(X509Certificate);

            //=========================================================================================================================  

            doc.Save(filePath);*/
           
            // insert xml in database 
            /*tei_company_pk = dt.Rows[0]["keycompany"].ToString();
            OracleBlob tempLob;
            OracleTransaction tx;

            Stream theStream = File.Open(filePath, FileMode.Open); //FileInput.PostedFile.InputStream;
            Byte[] buff = new Byte[(int)theStream.Length];
            theStream.Read(buff, 0, (int)theStream.Length);

            tx = connection.BeginTransaction();
            command = connection.CreateCommand();
            command.Transaction = tx;

            command.CommandText = "declare xx blob; begin dbms_lob.createtemporary(xx, false, 0); :tempblob := xx; end;";
            command.Parameters.Add("tempblob", OracleDbType.Blob).Direction = ParameterDirection.Output;
            command.Parameters["tempblob"].Size = (int)theStream.Length;
            command.ExecuteNonQuery();

            tempLob = (OracleBlob)command.Parameters[0].Value;
            tempLob.BeginChunkWrite();//.BeginBatch(OracleBlobOpenMode.ReadWrite);
            tempLob.Write(buff, 0, buff.Length);
            tempLob.EndChunkWrite();// EndBatch();

            command.Parameters.Clear();

            command.CommandText = "es_insert_image";

            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.Add("p_table_name", OracleDbType.Varchar2, 100);
            command.Parameters["p_table_name"].Value = "TEI_EINVOICE_XML_V2";

            command.Parameters.Add("p_master_pk", OracleDbType.Varchar2, 20);
            command.Parameters["p_master_pk"].Value = tei_company_pk;

            command.Parameters.Add("p_tc_fsbinary_pk", OracleDbType.Varchar2, 100);
            command.Parameters["p_tc_fsbinary_pk"].Value = tei_einvoice_m_pk;

            command.Parameters.Add("p_data", OracleDbType.Blob);
            command.Parameters["p_data"].Value = tempLob;
            command.Parameters["p_data"].Size = (int)theStream.Length;

            command.Parameters.Add("p_filename", OracleDbType.Varchar2, 100);
            command.Parameters["p_filename"].Value = dt.Rows[0]["templateCode"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceSerialNo"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceNumber"] + ".xml";

            command.Parameters.Add("p_filesize", OracleDbType.Varchar2, 20);
            command.Parameters["p_filesize"].Value = Math.Round((double)theStream.Length / 1024, 2);

            command.Parameters.Add("p_contenttype", OracleDbType.Varchar2, 100);
            command.Parameters["p_contenttype"].Value = "image/Xml";//fileStream.GetType();// FileInput.PostedFile.ContentType;

            command.Parameters.Add("p_crt_by", OracleDbType.Varchar2, 10);
            command.Parameters["p_crt_by"].Value = "useroutside";

            command.Parameters.Add("p_rtn_pk", OracleDbType.Varchar2, 200);
            command.Parameters["p_rtn_pk"].Direction = ParameterDirection.Output;

            command.ExecuteNonQuery();

            tx.Commit();*/
            connection.Close();
            connection.Dispose();
            return htmlStr.ToString() + "|" + dt.Rows[0]["templateCode"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceSerialNo"].ToString().Replace("/", "") + "_" + dt.Rows[0]["InvoiceNumber"];

        }

        public static String NumberToTextVN(decimal total)
        {
            try
            {
                string rs = "";
                total = Math.Round(total, 0);
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