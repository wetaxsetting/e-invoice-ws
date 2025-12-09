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
//using Oracle.ManagedDataAccess.Client;

namespace EInvoice.Company
{
    public class DaeWon_New
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

            int pos = 10, pos_lv = 20, v_count = 0, count_page = 0, count_page_v = 0, r = 0, x = 0;

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
            // read_amount = dt.Rows[0]["TotalAmountInWord"].ToString();

            // if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            // {
            //     read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmountInWord"].ToString()));
            // }
            // else
            // {
            //     read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            // }
            // read_prive = read_prive.Replace(",", "").Replace("TRừ", "Trừ");

            //read_prive = read_prive.ToString().Substring(0, 2) + read_prive.ToString().Substring(2, read_prive.Length - 2).ToLower().Replace("mỹ", "Mỹ");

            //read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower() + '.';

            read_prive = dt.Rows[0]["amount_word_vie"].ToString();

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

            htmlStr.Append("<!DOCTYPE html PUBLIC ' -//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																													 \n");
            htmlStr.Append("<html>                                                                                                                                                                                                                   \n");
            htmlStr.Append("<head>                                                                                                                                                                                                                   \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html;charset = UTF-8'>                                                                                                                                                      \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("<script type='text/javascript'                                                                                                                                                                                           \n");
            htmlStr.Append("	src='${ pageContext.request.contextPath}/system/syscommand.js'></script>                                                                                                                                              \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                                                                          \n");
            htmlStr.Append("<!-- Normalize or reset CSS with your favorite library -->                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("<!-- Load paper.css for happy printing -->                                                                                                                                                                               \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("<!-- Set page size here: A5, A4 or A3 -->                                                                                                                                                                                \n");
            htmlStr.Append("<!-- Set also 'landscape' if you need -->                                                                                                                                                                                \n");
            htmlStr.Append("<style>                                                                                                                                                                                                                  \n");
            htmlStr.Append("@page {                                                                                                                                                                                                                  \n");
            htmlStr.Append("	size: A4                                                                                                                                                                                                             \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("</style>                                                                                                                                                                                                                 \n");
            htmlStr.Append("<style>                                                                                                                                                                                                                  \n");
            htmlStr.Append("/*body   { font-family: serif }                                                                                                                                                                                          \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                                                                      \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                                                                      \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                                                                         \n");
            htmlStr.Append("    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                                                                                                         \n");
            htmlStr.Append("    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                                                                                                         \n");
            htmlStr.Append("    li     { font-size: 11pt; line-height: 5mm }                                                                                                                                                                         \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                                                                \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                                                                        \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                                                                         \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                                                              \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                                                                     \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                                                                    \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                                                                     \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                                                                        \n");
            htmlStr.Append("body {                                                                                                                                                                                                                   \n");
            htmlStr.Append("	color: blue;                                                                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 100%;                                                                                                                                                                                                     \n");
            htmlStr.Append("	background-image: url('assets/Solution.jpg');                                                                                                                                                                        \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("h1 {                                                                                                                                                                                                                     \n");
            htmlStr.Append("	color: #00FF00;                                                                                                                                                                                                      \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("p {                                                                                                                                                                                                                      \n");
            htmlStr.Append("	color: rgb(0, 0, 255)                                                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("headline1 {                                                                                                                                                                                                              \n");
            htmlStr.Append("	background-image: url(assets/Solution.jpg);                                                                                                                                                                          \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                                        \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                                       \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                                   \n");
            htmlStr.Append("	margin-bottom: 50px;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("headline2 {                                                                                                                                                                                                              \n");
            htmlStr.Append("	background-image: url(images/newsletter_headline2.gif);                                                                                                                                                              \n");
            htmlStr.Append("	background-repeat: no-repeat;                                                                                                                                                                                        \n");
            htmlStr.Append("	background-position: left top;                                                                                                                                                                                       \n");
            htmlStr.Append("	padding-top: 68px;                                                                                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append("<!--                                                                                                                                                                                                                     \n");
            htmlStr.Append("table {                                                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-displayed-decimal-separator: '\\, ';                                                                                                                                                                               \n");
            htmlStr.Append("	mso-displayed-thousand-separator: '\\.';                                                                                                                                                                              \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font525524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font625524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 14.5pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font725524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font825524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	color: #0066CC;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font925524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font1025524 {                                                                                                                                                                                                           \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font1125524 {                                                                                                                                                                                                           \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font1225524 {                                                                                                                                                                                                           \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font1325524 {                                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size: 20.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font1425524 {                                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size: 18.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font1525524 {                                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".font1625524 {                                                                                                                                                                                                           \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl6325524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl6425524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 2.0pt double windowtext;                                                                                                                                                                              \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl6525524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 2.0pt double windowtext;                                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl6625524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 2.0pt double windowtext;                                                                                                                                                                              \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl6725524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 2.0pt double windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl6825524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 2.0pt double windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl6925524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 2.0pt double windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7025524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7125524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7225524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7325524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7425524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7525524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 14.4pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7625524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7725524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7825524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl7925524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8025524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8125524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8225524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8325524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8425524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border: 1.0pt solid windowtext;                                                                                                                                                                                       \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8525524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8625524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8725524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8825524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl8925524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9025524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9125524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9225524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9325524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9425524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9525524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9625524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9725524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9825524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl9925524 {                                                                                                                                                                                                             \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl99255241 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append(".xl10025524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10125524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10225524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10325524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                                                                 \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10425524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10525524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10625524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10725524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #FFC000;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10825524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #FFC000;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl10925524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #FFC000;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11025524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11125524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11225524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: black;                                                                                                                                                                                                        \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                                                                      \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11325524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11425524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11525524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11625524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11725524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11825524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl11925524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12025524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12125524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12225524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12325524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12425524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12525524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12625524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12725524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12825524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl12925524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13025524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13125524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13225524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13325524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                                   \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13425524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13525524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13625524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13725524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13825524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl13925524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14025524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14125524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14225524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14325524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14425524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14525524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14625524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14725524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 12.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1pt dotted windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1pt dotted windowtext;                                                                                                                                                                               \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14825524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 11.25pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: top;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl14925524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                                                                   \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15025524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format:                                                                                                                                                                                                   \n");
            htmlStr.Append("		'_\\(* \\#\\,\\#\\#0\\.00_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\.00\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                                                                                            \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15125524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	font-size: 15pt;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15225524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15325524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15425524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15525524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 13.75pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                   \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15625524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 18.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15725524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15825524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: 1.0pt solid windowtext;                                                                                                                                                                                  \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl15925524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl16025524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl16125524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: 1.0pt solid windowtext;                                                                                                                                                                                \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                                                                   \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl16225524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl16325524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: red;                                                                                                                                                                                                          \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                                                         \n");
            htmlStr.Append(".xl16425524 {                                                                                                                                                                                                            \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                                                                        \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                                                                 \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-size: 15.0pt;                                                                                                                                                                                                   \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                                                                    \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                                                                  \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                                                               \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                                                               \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                                                                          \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                                                                  \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                                                              \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                                                                    \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                                                                  \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                                                                 \n");
            htmlStr.Append("	border-left: 1.0pt solid windowtext;                                                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                                                                         \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                                                                   \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                                                                        \n");
            htmlStr.Append("-->                                                                                                                                                                                                                      \n");
            htmlStr.Append("</style>                                                                                                                                                                                                                 \n");
            htmlStr.Append("</head>                                                                                                                                                                                                                  \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                                                                        \n");

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

            double v_totalHeightLastPage = 303.5;// 373.5.5

            double v_totalHeightPage = 580;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 505;// 540;

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

                htmlStr.Append("	<table border=0 cellpadding=0 cellspacing=0 width=696 class=xl6325524                                                                                                                                                \n");
                htmlStr.Append("		style='border-collapse: collapse; table-layout: fixed; width: 784.5pt'>                                                                                                                                            \n");
                htmlStr.Append("		<col class=xl6325524 width=38                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1336; width: 42pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=81                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2872; width: 91.5pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=84                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2986; width: 49.5pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=62                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2218; width: 108pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=84                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2986; width: 102pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=41                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1450; width: 46.5pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=74                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2616; width: 82.5pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=106                                                                                                                                                                                   \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 3783; width: 120pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=52                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 1848; width: 58.5pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=62                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 2218; width: 70.5pt'>                                                                                                                                         \n");
                htmlStr.Append("		<col class=xl6325524 width=12                                                                                                                                                                                    \n");
                htmlStr.Append("			style='mso-width-source: userset; mso-width-alt: 426; width: 13.5pt'>                                                                                                                                           \n");
                htmlStr.Append("			<tr height=22 style='mso-height-source: userset; height: 21.1875pt'>                                                                                                                                           \n");
                htmlStr.Append("				<td colspan=1 height=22 width=696  class=xl15625524                                                                                                                                                                     \n");
                htmlStr.Append("					style='border-right: none; height: 21.1875pt; width: 523pt'                                                                                                                               \n");
                htmlStr.Append("					align=left valign=top><![if !vml]><span                                                                                                                                                              \n");
                htmlStr.Append("					style='mso-ignore: vglayout; position: absolute; z-index: 1;margin-left: 15px; margin-bottom: 14px; width:120px; height: 80px'><img                                                                    \n");
                htmlStr.Append("						width=120 height=80                                                                                                                                                                               \n");
                htmlStr.Append("						src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\DAEWON_0001.png'                                                                                                                           \n");
                htmlStr.Append("						alt=logo v:shapes='Picture_x0020_1'></span> <![endif]></td>                                                                                                                                                                                             \n");
                htmlStr.Append("				<td colspan=10 height=22 width=696  class=xl15625524                                                                                                                                                                     \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black;border-left: none; height: 21.1875pt; width: 523pt'                                                                                                                               \n");
                htmlStr.Append("					align=left valign=top><span                                                                                                                                      \n");
                htmlStr.Append("					style='mso-ignore: vglayout2'>                                                                                                                                                                       \n");
                htmlStr.Append("						<table cellpadding=0 cellspacing=0>                                                                                                                                                              \n");
                htmlStr.Append("							<tr>                                                                                                                                                                                         \n");
                htmlStr.Append("								<td colspan=11 height=22  width=696                                                                                                                                      \n");
                htmlStr.Append("									style='border-right: none; height: 21.1875pt; width: 523pt;text-align:right'>" + dt.Rows[0]["Seller_Name"] + "<span style='mso-spacerun: yes'></span></td>                                                                                                         \n");
                htmlStr.Append("							</tr>                                                                                                                                                                                        \n");
                htmlStr.Append("						</table>                                                                                                                                                                                         \n");
                htmlStr.Append("				</span></td>                                                                                                                                                                                             \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=21 style='height: 19.5pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td colspan=11 height=21 class=xl15925524                                                                                                                                                                \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black; height: 19.5pt'>" + dt.Rows[0]["Seller_Address"] + "</td>                                                                                                                      \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=21 style='height: 19.5pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td colspan=11 height=21 class=xl15925524                                                                                                                                                                \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black; height: 19.5pt'>Tel: " + dt.Rows[0]["Seller_Tel"] + "                                                                                                                          \n");
                htmlStr.Append("					* Fax: " + dt.Rows[0]["Seller_Fax"] + "</td>                                                                                                                                                                           \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=21 style='height: 19.5pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td colspan=11 height=21 class=xl15925524                                                                                                                                                                \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black; height: 19.5pt'>Mã                                                                                                                                           \n");
                htmlStr.Append("					s&#7889; thu&#7871; : " + dt.Rows[0]["Seller_TaxCode"] + "</td>                                                                                                                                                        \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=12 style='mso-height-source: userset; height: 2.5pt'>                                                                                                                                             \n");
                htmlStr.Append("				<td height=12 class=xl6425524 style='height: 2.5pt'>&nbsp;</td>                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6525524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl6625524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=33 style='height: 30.75pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td colspan=7 height=33 class=xl16225524 style='height: 30.75pt'><font                                                                                                                                    \n");
                htmlStr.Append("					class='font1325524'>HÓA &#272;&#416;N GTGT/ </font><font                                                                                                                                             \n");
                htmlStr.Append("					class='font1425524'>VAT INVOICE</font></td>                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl7025524 colspan=4                                                                                                                                                                            \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black'><span                                                                                                                                      \n");
                htmlStr.Append("					style='mso-spacerun: yes'>  </span><font class='font625524'></font><font class='font525524'></font></td>                                                                                                                                      \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=21 style='height: 19.5pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td colspan=7 height=21 class=xl11725524 style='height: 19.5pt'>                                                                                                                                         \n");
                htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
                htmlStr.Append("				<td class=xl7025524 colspan=3>Ký hi&#7879;u / <font                                                                                                                                                      \n");
                htmlStr.Append("					class='font625524'>Serial no</font><font class='font525524'>:                                                                                                                                        \n");
                htmlStr.Append("				</font><font class='font925524'>" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</font></td>                                                                                                                                                 \n");
                htmlStr.Append("				<td class=xl7125524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=30 style='mso-height-source: userset; height: 28.125pt'>                                                                                                                                            \n");
                htmlStr.Append("				<td colspan=7 height=30 class=xl16425524 style='height: 28.125pt'>Ngày                                                                                                                                     \n");
                htmlStr.Append("					hóa &#273;&#417;n<span style='mso-spacerun: yes'>  </span>/ <font                                                                                                                                    \n");
                htmlStr.Append("					class='font625524'>Invoice date:<span                                                                                                                                                                \n");
                htmlStr.Append("						style='mso-spacerun: yes'> </span></font> <font style='font-weight: 700'>" + dt.Rows[0]["invoiceissueddate_dd"] + " /" + dt.Rows[0]["invoiceissueddate_mm"] + " /" + dt.Rows[0]["invoiceissueddate_yyyy"] + " </font>                                                                               \n");
                htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
                htmlStr.Append("				<td class=xl7025524 colspan=3>S&#7889; / <font                                                                                                                                                           \n");
                htmlStr.Append("					class='font625524'>Invoice no</font><font class='font525524'>:                                                                                                                                       \n");
                htmlStr.Append("				</font><font class='font1625524'>" + dt.Rows[0]["InvoiceNumber"] + " </font></td>                                                                                                                                              \n");
                htmlStr.Append("				<td class=xl7125524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=21 style='height: 19.5pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td height=21 class=xl7225524 style='height: 19.5pt'>&nbsp;</td>                                                                                                                                         \n");
                htmlStr.Append("				<td class=xl7325524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl7325524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl7325524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl7325524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl7325524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl7325524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("				<td class=xl7325524 colspan=3>" + v_titlePageNumber +"</td>                                                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl7425524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr class=xl7025524 height=28                                                                                                                                                                                \n");
                htmlStr.Append("				style='mso-height-source: userset; height: 23.0pt'>                                                                                                                                                     \n");
                htmlStr.Append("				<td height=28 class=xl7525524 colspan=10 style='height: 23.0pt'>&nbsp;H&#7885;                                                                                                                           \n");
                htmlStr.Append("					tên ng&#432;&#7901;i mua hàng / <font class='font625524'>Customer's                                                                                                                                  \n");
                htmlStr.Append("						name</font><font class='font525524'> :</font>&nbsp;" + dt.Rows[0]["buyer"] + "                                                                                                                                \n");
                htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
                //htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                //htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                //htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                //htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                //htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl7625524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr class=xl7025524 height=28                                                                                                                                                                                \n");
                htmlStr.Append("				style='mso-height-source: userset; height: 23.0pt'>                                                                                                                                                     \n");
                htmlStr.Append("				<td height=28 class=xl7525524 colspan=3 style='height: 23.0pt'>&nbsp;Tên                                                                                                                               \n");
                htmlStr.Append("					&#273;&#417;n v&#7883; / <font class='font625524'>Company</font><font                                                                                                                                \n");
                htmlStr.Append("					class='font525524'> :<span style='mso-spacerun: yes'> </span></font>                                                                                                                    \n");
                htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
                htmlStr.Append("				<td class=xl7025524 colspan=7 >" + dt.Rows[0]["buyerlegalname"] + "</td>                                                                                                                                               \n");
                htmlStr.Append("				<td class=xl7625524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr class=xl7025524 height=28                                                                                                                                                                                \n");
                htmlStr.Append("				style='mso-height-source: userset; height: 23.0pt'>                                                                                                                                                     \n");
                htmlStr.Append("				<td height=28 class=xl7525524 colspan=10 style='height: 23.0pt'>&nbsp;&#272;&#7883;a                                                                                                                    \n");
                htmlStr.Append("					ch&#7881; / <font class='font625524'>&nbsp;Address</font><font                                                                                                                                       \n");
                htmlStr.Append("					class='font525524'> :<span style='mso-spacerun: yes'> </span></font>" + dt.Rows[0]["buyeraddress"] + "                                                                                                                       \n");
                htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
                htmlStr.Append("				<td class=xl7625524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr class=xl7025524 height=25                                                                                                                                                                                \n");
                htmlStr.Append("				style='mso-height-source: userset; height: 23.0pt'>                                                                                                                                                     \n");
                htmlStr.Append("				<td height=25 class=xl7525524 colspan=5 style='height: 23.0pt'>&nbsp;Hình                                                                                                                               \n");
                htmlStr.Append("					th&#7913;c thanh toán / <font class='font625524'>Term of                                                                                                                                             \n");
                htmlStr.Append("						payment</font><font class='font525524'> :<span                                                                                                                                                   \n");
                htmlStr.Append("						style='mso-spacerun: yes'>           </span></font> " + dt.Rows[0]["PaymentMethodCK"] + "                                                                                                                            \n");
                htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
                htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl7025524 colspan=2>Mã s&#7889; thu&#7871; / <font                                                                                                                                             \n");
                htmlStr.Append("					class='font625524'>Tax code</font><font class='font525524'>                                                                                                                                          \n");
                htmlStr.Append("						:<span style='mso-spacerun: yes'> </span>" + dt.Rows[0]["BuyerTaxCode"] + "                                                                                                                                             \n");
                htmlStr.Append("				</font></td>                                                                                                                                                                                             \n");
                htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl7025524></td>                                                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl7625524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr class=xl7025524 height=21 style='height: 23.0pt'>                                                                                                                                                       \n");
                htmlStr.Append("				<td height=21 class=xl7525524 colspan=10 style='height: 23.0pt'>&nbsp;Giao                                                                                                                               \n");
                htmlStr.Append("					hàng t&#7841;i kho cty :<span style='mso-spacerun: yes'> </span> " + dt.Rows[0]["attribute_01"] + "                                                                                                                      \n");
                htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
                htmlStr.Append("				<td class=xl7625524>&nbsp;</td>                                                                                                                                                                          \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=13                                                                                                                                                                                                \n");
                htmlStr.Append("				style='mso-height-source: userset; height: 2.05pt; font-size: 2.5pt;'>                                                                                                                                   \n");
                htmlStr.Append("				<td height=13 class=xl7725524                                                                                                                                                                            \n");
                htmlStr.Append("					style='height: 2.05pt; font-size: 2.5pt;'>&nbsp;</td>                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl6325524 style='height: 2.05pt; font-size: 2.5pt;'></td>                                                                                                                                      \n");
                htmlStr.Append("				<td class=xl7125524 style='height: 2.05pt; font-size: 2.5pt;'>&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=21 style='height: 19.5pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td height=21 class=xl7825524 style='height: 19.5pt'>STT</td>                                                                                                                                            \n");
                htmlStr.Append("				<td colspan=4 class=xl7925524 style='border-left: none'>Tên                                                                                                                                              \n");
                htmlStr.Append("					hàng hóa, d&#7883;ch v&#7909;</td>                                                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl7925524 style='border-left: none'>&#272;VT</td>                                                                                                                                              \n");
                htmlStr.Append("				<td class=xl7925524 style='border-left: none'>S&#7889;                                                                                                                                                   \n");
                htmlStr.Append("					l&#432;&#7907;ng</td>                                                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl7925524 style='border-left: none'>&#272;&#417;n giá                                                                                                                                          \n");
                htmlStr.Append("					(UP)</td>                                                                                                                                                                                            \n");
                htmlStr.Append("				<td colspan=3 class=xl7925524                                                                                                                                                                            \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black; border-left: none'>Thành                                                                                                                                     \n");
                htmlStr.Append("					ti&#7873;n</td>                                                                                                                                                                                      \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr height=21 style='height: 19.5pt'>                                                                                                                                                                        \n");
                htmlStr.Append("				<td height=21 class=xl8025524 style='height: 19.5pt'>No</td>                                                                                                                                             \n");
                htmlStr.Append("				<td colspan=4 class=xl8125524 style='border-left: none'>Description</td>                                                                                                                                 \n");
                htmlStr.Append("				<td class=xl8125524 style='border-left: none'>Unit</td>                                                                                                                                                  \n");
                htmlStr.Append("				<td class=xl8125524 style='border-left: none'>Quantity</td>                                                                                                                                              \n");
                htmlStr.Append("				<td class=xl8125524 style='border-left: none'>" + dt.Rows[0]["CurrencyCodeUSD"] + " </td>                                                                                                                                            \n");
                htmlStr.Append("				<td colspan=3 class=xl8125524                                                                                                                                                                            \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black; border-left: none'>Amount</td>                                                                                                                               \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			<tr class=xl8525524 height=18 style='height: 17.25pt'>                                                                                                                                                        \n");
                htmlStr.Append("				<td height=18 class=xl8325524                                                                                                                                                                            \n");
                htmlStr.Append("					style='height: 17.25pt; border-top: none'>A</td>                                                                                                                                                      \n");
                htmlStr.Append("				<td colspan=4 class=xl8425524 style='border-left: none'>B</td>                                                                                                                                           \n");
                htmlStr.Append("				<td class=xl8425524 style='border-top: none; border-left: none'>C</td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl8425524 style='border-top: none; border-left: none'>1</td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl8425524 style='border-top: none; border-left: none'>2</td>                                                                                                                                   \n");
                htmlStr.Append("				<td colspan=3 class=xl8425524                                                                                                                                                                            \n");
                htmlStr.Append("					style='border-right: 1.0pt solid black; border-left: none'>3 =                                                                                                                                       \n");
                htmlStr.Append("					1 x 2</td>                                                                                                                                                                                           \n");
                htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
                htmlStr.Append("			                                                                                                                                                                                                             \n");


                v_rowHeight = "38.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 38;

                v_rowHeightLast = "35.0pt";// "23.5pt";
                v_rowHeightLastNumber = 35;// 23.5;
                v_rowHeightEmptyLast = "35.0pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 40)
                    {
                        v_rowHeight = "35.0pt"; //"26.5pt";    
                        v_rowHeightLast = "33.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 33;//27.5;
                        v_rowHeightEmptyLast = "35.0pt"; //"23.0pt";
                        vlongItemName = true;
                    }
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

                        htmlStr.Append("					 	                                                                                                                                                                                                 \n");
                        htmlStr.Append("									<tr class=xl10025524 height=29                                                                                                                                                       \n");
                        htmlStr.Append("										style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                             \n");
                        htmlStr.Append("										<td height=29 class=xl9925524 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                       \n");
                        htmlStr.Append("										<td colspan=4 class=xl15125524 style='border-left: none'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                    \n");
                        htmlStr.Append("										<td class=xl10125524 style='border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                                              \n");
                        htmlStr.Append("										<td class=xl10225524 style='border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                              \n");
                        htmlStr.Append("										<td class=xl10225524 style='border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                              \n");
                        htmlStr.Append("										<td colspan=3 class=xl14525524                                                                                                                                                   \n");
                        htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                              \n");
                        htmlStr.Append("									</tr>                                                                                                                                                                                \n");
                        htmlStr.Append("                                                                                                                                                                                                                         \n");
                        htmlStr.Append("			                                                                                                                                                                                                             \n");
                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {

                            htmlStr.Append("									<tr class=xl10025524 height=29                                                                                                                                                       \n");
                            htmlStr.Append("										style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                             \n");
                            htmlStr.Append("										<td height=29 class=xl9925524                                                                                                                                                    \n");
                            htmlStr.Append("											style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                               \n");
                            htmlStr.Append("										<td colspan=4 class=xl14225524                                                                                                                                                   \n");
                            htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                               \n");
                            htmlStr.Append("										<td class=xl10125524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td colspan=3 class=xl14525524                                                                                                                                                   \n");
                            htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("									</tr>                                                                                                                                                                                \n");

                       

                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("									<tr class=xl10025524 height=29                                                                                                                                                       \n");
                                htmlStr.Append("										style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                             \n");
                                htmlStr.Append("										<td height=29 class=xl9925524                                                                                                                                                    \n");
                                htmlStr.Append("											style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                               \n");
                                htmlStr.Append("										<td colspan=4 class=xl14225524                                                                                                                                                   \n");
                                htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                               \n");
                                htmlStr.Append("										<td class=xl10125524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                            \n");
                                htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                            \n");
                                htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                            \n");
                                htmlStr.Append("										<td colspan=3 class=xl14525524                                                                                                                                                   \n");
                                htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                              \n");
                                htmlStr.Append("									</tr>                                                                                                                                                                                \n");
                            }
                            else
                            {
                                htmlStr.Append("									<tr class=xl10025524 height=29                                                                                                                                                       \n");
                                htmlStr.Append("										style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                             \n");
                                htmlStr.Append("										<td height=29 class=xl9925524                                                                                                                                                    \n");
                                htmlStr.Append("											style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                               \n");
                                htmlStr.Append("										<td colspan=4 class=xl14225524                                                                                                                                                   \n");
                                htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                               \n");
                                htmlStr.Append("										<td class=xl10125524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                            \n");
                                htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                            \n");
                                htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                            \n");
                                htmlStr.Append("										<td colspan=3 class=xl14525524                                                                                                                                                   \n");
                                htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                              \n");
                                htmlStr.Append("									</tr>                                                                                                                                                                                \n");
                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                        htmlStr.Append("									<tr class=xl10025524 height=29                                                                                                                                                       \n");
                        htmlStr.Append("										style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                                             \n");
                        htmlStr.Append("										<td height=29 class=xl9925524                                                                                                                                                    \n");
                        htmlStr.Append("											style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                               \n");
                        htmlStr.Append("										<td colspan=4 class=xl14225524                                                                                                                                                   \n");
                        htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                               \n");
                        htmlStr.Append("										<td class=xl10125524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][1] + "&nbsp;</td>                                                                                            \n");
                        htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                            \n");
                        htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                            \n");
                        htmlStr.Append("										<td colspan=3 class=xl14525524                                                                                                                                                   \n");
                        htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                              \n");
                        htmlStr.Append("									</tr>                                                                                                                                                                                \n");

                    }
                    v_index++;
                } //for dtR

                v_spacePerPage = 0;
                if (k < v_countNumberOfPages - 1 && page[k] < rowsPerPage)
                {
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        v_spacePerPage += v_rowHeightNumber;//  v_totalHeightPage;
                    }
                }
                else if (k < v_countNumberOfPages - 1 && page[k] == rowsPerPage)
                {
                    v_spacePerPage = 36;
                }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("									<tr class=xl10025524 height=29                                                                                                                                                       \n");
                            htmlStr.Append("										style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                                                             \n");
                            htmlStr.Append("										<td height=29 class=xl9925524                                                                                                                                                    \n");
                            htmlStr.Append("											style='height: " + v_rowHeightEmptyLast + "; border-top: none'>&nbsp;</td>                                                                                                               \n");
                            htmlStr.Append("										<td colspan=4 class=xl14225524                                                                                                                                                   \n");
                            htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;</td>                                                                                               \n");
                            htmlStr.Append("										<td class=xl10125524 style='border-top: none; border-left: none'>&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td colspan=3 class=xl14525524                                                                                                                                                   \n");
                            htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("									</tr>                                                                                                                                                                                \n");

                        }
                        else
                        {
                            htmlStr.Append("									<tr class=xl10025524 height=29                                                                                                                                                       \n");
                            htmlStr.Append("										style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                                                                                             \n");
                            htmlStr.Append("										<td height=29 class=xl9925524                                                                                                                                                    \n");
                            htmlStr.Append("											style='height: " + v_rowHeightEmptyLast + "; border-top: none'>&nbsp;</td>                                                                                                               \n");
                            htmlStr.Append("										<td colspan=4 class=xl14225524                                                                                                                                                   \n");
                            htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;</td>                                                                                               \n");
                            htmlStr.Append("										<td class=xl10125524 style='border-top: none; border-left: none'>&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td class=xl10225524 style='border-top: none; border-left: none'>&nbsp;</td>                                                                                            \n");
                            htmlStr.Append("										<td colspan=3 class=xl14525524                                                                                                                                                   \n");
                            htmlStr.Append("											style='border-right: 1.0pt solid black; border-left: none'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("									</tr>                                                                                                                                                                                \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: " + (v_spacePerPage).ToString() + "pt'>                                                                                                                                               \n");
                    htmlStr.Append("			<td colspan=11 height=22 class=xl11325524                                                                                                                                                                    \n");
                    htmlStr.Append("				style='border-right: 1.0pt solid black; height: " + (v_spacePerPage).ToString() + "pt'>(C&#7847;n                                                                                                                                      \n");
                    htmlStr.Append("				ki&#7875;m tra, &#273;&#7889;i chi&#7871;u khi l&#7853;p, giao,                                                                                                                                          \n");
                    htmlStr.Append("				nh&#7853;n hóa &#273;&#417;n)</td>                                                                                                                                                                       \n");
                    htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
                    htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                                            \n");
                    htmlStr.Append("			<td colspan=11 height=21 class=xl11625524 style='height: 19.5pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                                                                     \n");
                    htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
                    htmlStr.Append("		<tr height=21 style='height: 30.5pt'>                                                                                                                                                                            \n");
                    htmlStr.Append("			<td colspan=11 height=21  style='height: 30.5pt'></td>                                                                                                                                                     \n");
                    htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
                    htmlStr.Append("		<![if supportMisalignedColumns]>                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=0 style='display: none'>                                                                                                                                                                              \n");
                    htmlStr.Append("			<td width=38 style='width: 35pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=81 style='width: 76.25pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=84 style='width: 78.75pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=62 style='width: 58.75pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=84 style='width: 78.75pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=41 style='width: 38.75pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=74 style='width: 68.75pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=106 style='width: 100pt'></td>                                                                                                                                                                      \n");
                    htmlStr.Append("			<td width=52 style='width: 48.75pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=62 style='width: 58.75pt'></td>                                                                                                                                                                       \n");
                    htmlStr.Append("			<td width=12 style='width: 11.25pt'></td>                                                                                                                                                                        \n");
                    htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
                    htmlStr.Append("		<![endif]>                                                                                                                                                                                                       \n");
                    htmlStr.Append("	</table>                                                                                                                                                                                                             \n");
                    htmlStr.Append("	                                                                                                                                                                                                                     \n");

                }


            }// for k                                                                                                                             
            htmlStr.Append("		<tr height=20 style='mso-height-source: userset; height: 23.0pt'>                                                                                                                                                \n");
            htmlStr.Append("			<td height=20 class=xl8625524 style='height: 23.0pt'>&nbsp;</td>                                                                                                                                             \n");
            htmlStr.Append("			<td colspan=4 class=xl14825524 width=311                                                                                                                                                                     \n");
            htmlStr.Append("				style='border-left: none; width: 234pt'>&nbsp;</td>                                                                                                                                                      \n");
            htmlStr.Append("			<td class=xl8725524 style='border-left: none'>&nbsp;</td>                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl8725524 style='border-left: none'>&nbsp;</td>                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl10325524 style='border-left: none'>Tổng cộng " + dt.Rows[0]["CurrencyCodeUSD"] + "                                                                                                                                           \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=3 class=xl14925524                                                                                                                                                                               \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>" + dt.Rows[0]["netamount_display"] + "&nbsp;</td>                                                                                                                         \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=5 height=21 class=xl13025524 style='height: 19.5pt'>&nbsp;T&#7927;                                                                                                                               \n");
            htmlStr.Append("				giá / <font class='font625524'>Exchange rate</font><font                                                                                                                                                 \n");
            htmlStr.Append("				class='font525524'> :<span style='mso-spacerun: yes'> </span></font>" + dt.Rows[0]["exchangerate_no"] + "                                                                                                                         \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=3 class=xl13225524 style='border-right: 1.0pt solid black'>C&#7897;ng                                                                                                                             \n");
            htmlStr.Append("				ti&#7873;n hàng / <font class='font625524'>Total Amount</font><font                                                                                                                                      \n");
            htmlStr.Append("				class='font525524'> :<span style='mso-spacerun: yes'>  </span></font>                                                                                                                                    \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=3 class=xl13725524                                                                                                                                                                               \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>" + dt.Rows[0]["netamount_display"] + "&nbsp;</td>                                                                                                                         \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr class=xl7025524 height=21 style='height: 19.5pt'>                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=5 height=21 class=xl13025524 style='height: 19.5pt'>&nbsp;Thu&#7871;                                                                                                                             \n");
            htmlStr.Append("				su&#7845;t GTGT / <font class='font625524'>VAT rate&nbsp;</font><font                                                                                                                                    \n");
            htmlStr.Append("				class='font525524'> :&nbsp;<span style='mso-spacerun: yes'>  </span>&nbsp;" + dt.Rows[0]["TaxRate"] + "                                                                                                                    \n");
            htmlStr.Append("			</font>                                                                                                                                                                                                      \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=3 class=xl13225524 style='border-right: 1.0pt solid black'>&nbsp;Ti&#7873;n                                                                                                                       \n");
            htmlStr.Append("				thu&#7871; GTGT /<font class='font625524'> VAT Amount </font><font                                                                                                                                       \n");
            htmlStr.Append("				class='font525524'>:<span style='mso-spacerun: yes'> &nbsp;                                                                                                                                              \n");
            htmlStr.Append("				</span></font>                                                                                                                                                                                           \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=3 class=xl13425524                                                                                                                                                                               \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>" + amout_vat + "&nbsp;</td>                                                                                                                       \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr class=xl7025524 height=21 style='height: 19.5pt'>                                                                                                                                                            \n");
            htmlStr.Append("			<td height=21 class=xl8825524                                                                                                                                                                                \n");
            htmlStr.Append("				style='height: 19.5pt; border-top: none'>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl8925524 style='border-top: none'>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl8925524 style='border-top: none'>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append("			<td class=xl8925524 style='border-top: none'>&nbsp;</td>                                                                                                                                                     \n");
            htmlStr.Append("			<td colspan=4 class=xl13225524 style='border-right: 1.0pt solid black'>T&#7893;ng                                                                                                                             \n");
            htmlStr.Append("				c&#7897;ng ti&#7873;n thanh toán / <font class='font625524'>Total                                                                                                                                        \n");
            htmlStr.Append("					Amount</font><font class='font525524'> :<span                                                                                                                                                        \n");
            htmlStr.Append("					style='mso-spacerun: yes'>  </span></font>                                                                                                                                                           \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=3 class=xl13725524                                                                                                                                                                               \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; border-left: none'>" + dt.Rows[0]["totalamount_display"] + "&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=33 style='mso-height-source: userset; height: 31.3125pt'>                                                                                                                                               \n");
            htmlStr.Append("			<td height=33 class=xl9025524 colspan=2 style='height: 31.3125pt'>&nbsp;S&#7889;                                                                                                                               \n");
            htmlStr.Append("				ti&#7873;n b&#7857;ng ch&#7919; :</td>                                                                                                                                                                   \n");
            htmlStr.Append("			<td colspan=9 class=xl14025524                                                                                                                                                                               \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black'>&nbsp;&nbsp;" + read_prive + "&nbsp;                                                                                                                                    \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=33 style='mso-height-source: userset; height: 25.05pt'>														\n");
            htmlStr.Append("			<td height=33 class=xl9225524 colspan=3																				\n");
            htmlStr.Append("				style='height: 25.05pt;'>&nbsp;(Total																			\n");
            htmlStr.Append("				amount in words):																								\n");
            htmlStr.Append("			</td>																												\n");
            htmlStr.Append("			<td colspan=8 class=xl12525524																						\n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; border-bottom: none;'> <font											\n");
            htmlStr.Append("				style='font-size: 15pt; font-weight: 400; font-style: italic;'>" + read_en + "</font>&nbsp;</td>				\n");
            htmlStr.Append("		</tr>																													\n");
            htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=4 height=21 class=xl11925524 style='height: 19.5pt'>Ng&#432;&#7901;i                                                                                                                             \n");
            htmlStr.Append("				mua hàng / <font class='font625524'>Customer</font>                                                                                                                                                      \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td colspan=2 class=xl12125524></td>                                                                                                                                                                         \n");
            htmlStr.Append("			<td colspan=4 class=xl12425524                                                                                                                                                                               \n");
            htmlStr.Append("				style='border-top: 1pt dotted windowtext;'>Th&#7911;                                                                                                                                                    \n");
            htmlStr.Append("				tr&#432;&#7903;ng &#273;&#417;n v&#7883; /<font class='font625524'>                                                                                                                                      \n");
            htmlStr.Append("					Manager</font>                                                                                                                                                                                       \n");
            htmlStr.Append("			</td>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<td class=xl7125524 style='border-top: 1pt dotted windowtext;'>&nbsp;</td>                                                                                                                                  \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=4 height=21 class=xl12225524 style='height: 19.5pt'>(Ký,                                                                                                                                         \n");
            htmlStr.Append("				ghi rõ h&#7885; tên)</td>                                                                                                                                                                                \n");
            htmlStr.Append("			<td colspan=2 class=xl12425524><span style='mso-spacerun: yes'> </span></td>                                                                                                                                 \n");
            htmlStr.Append("			<td colspan=4 class=xl12425524>Ký, ghi rõ h&#7885; tên<span                                                                                                                                                  \n");
            htmlStr.Append("				style='mso-spacerun: yes'> </span></td>                                                                                                                                                                  \n");
            htmlStr.Append("			<td class=xl7125524>&nbsp;</td>                                                                                                                                                                              \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=4 height=21 class=xl12825524 style='height: 19.5pt'>(Sign                                                                                                                                        \n");
            htmlStr.Append("				&amp; full name)</td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td colspan=2 class=xl12725524></td>                                                                                                                                                                         \n");
            htmlStr.Append("			<td colspan=4 class=xl12725524>(Signature, full name)</td>                                                                                                                                                   \n");
            htmlStr.Append("			<td class=xl9425524>&nbsp;</td>                                                                                                                                                                              \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                                            \n");
            htmlStr.Append("			<td height=21 class=xl7725524 style='height: 19.5pt'>&nbsp;</td>                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td class=xl7125524>&nbsp;</td>                                                                                                                                                                              \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=20 style='mso-height-source: userset; height: 23.0pt'>                                                                                                                                                \n");
            htmlStr.Append("			<td height=20 class=xl7725524 style='height: 23.0pt'>&nbsp;</td>                                                                                                                                             \n");
            htmlStr.Append("			<td class=xl6325524 colspan=4>" + dt.Rows[0]["attribute_02"] + " </td>                                                                                                                                                           \n");
            htmlStr.Append("			<td class=xl6325524></td>                                                                                                                                                                                    \n");
            htmlStr.Append("			<td colspan=4 height=20 width=294 class=xl10425524                                                                                                                                                                           \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; height: 23.0pt; width: 221pt'                                                                                                                                     \n");
            htmlStr.Append("				align=left valign=top><![if !vml]><span                                                                                                                                                                  \n");
            htmlStr.Append("				style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 127px; margin-top: 15px; width: 46px; height: 37px'><img                                                                       \n");
            htmlStr.Append("					width=46 height=37                                                                                                                                                                                   \n");
            htmlStr.Append("					src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png'                                                                                                                                 \n");
            htmlStr.Append("					v:shapes='Picture_x0020_2'></span> <![endif]><span                                                                                                                                                   \n");
            htmlStr.Append("				style='mso-ignore: vglayout2'>                                                                                                                                                                           \n");
            htmlStr.Append("					<table cellpadding=0 cellspacing=0>                                                                                                                                                                  \n");
            htmlStr.Append("						<tr>                                                                                                                                                                                             \n");
            htmlStr.Append("							<td colspan=4 height=20  width=294                                                                                                                                           \n");
            htmlStr.Append("								style='border-right: ; height: 23.0pt; width: 221pt'>&nbsp;Signature                                                                                                                 \n");
            htmlStr.Append("								Valid<span style='mso-spacerun: yes'> </span>                                                                                                                                            \n");
            htmlStr.Append("							</td>                                                                                                                                                                                        \n");
            htmlStr.Append("						</tr>                                                                                                                                                                                            \n");
            htmlStr.Append("					</table>                                                                                                                                                                                             \n");
            htmlStr.Append("			</span></td>                                                                                                                                                                                                 \n");
            htmlStr.Append("				                                                                                                                                                                                                         \n");
            htmlStr.Append("			<td class=xl7125524>&nbsp;</td>                                                                                                                                                                              \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("			<tr height=28 style='mso-height-source: userset; height: 26.25pt'>                                                                                                                                            \n");
            htmlStr.Append("				<td height=28 class=xl9825524 colspan=2 style='height: 26.25pt'>&nbsp;                                                                                                                                       \n");
            htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td colspan=4 class=xl10725524 width=294                                                                                                                                                                 \n");
            htmlStr.Append("					style='border-right: 1.0pt solid black; width: 221pt'><font                                                                                                                                           \n");
            htmlStr.Append("					class='font1225524'>&nbsp;&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                                                                                             \n");
            htmlStr.Append("					class='font1525524'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                                                                                                                        \n");
            htmlStr.Append("				                                                                                                                                                                                                         \n");
            htmlStr.Append("				<td class=xl7125524>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<tr height=18 style='mso-height-source: userset; height: 16.875pt'>                                                                                                                                            \n");
            htmlStr.Append("				<td height=18 class=xl9525524 colspan=2 style='height: 16.875pt'>&nbsp;&nbsp;Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + " </td>                                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td colspan=4 class=xl11025524                                                                                                                                                                           \n");
            htmlStr.Append("					style='border-right: 1.0pt solid black'><font                                                                                                                                                         \n");
            htmlStr.Append("					class='font1125524'>&nbsp;Ngày ký</font><font class='font1025524'>:                                                                                                                                  \n");
            htmlStr.Append("						" + dt.Rows[0]["SignedDate"] + " </font></td>                                                                                                                                                                          \n");
            htmlStr.Append("				<td class=xl7125524>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
            htmlStr.Append("			<tr height=18 style='mso-height-source: userset; height: 17.8125pt'>                                                                                                                                           \n");
            htmlStr.Append("				<td height=18 class=xl9625524 colspan=5 style='height: 17.8125pt'>&nbsp;                                                                                                                                   \n");
            htmlStr.Append("					Tra c&#7913;u t&#7841;i Website: <font class='font725524'><span                                                                                                                                      \n");
            htmlStr.Append("						style='mso-spacerun: yes'> </span></font><font class='font825524'>" + dt.Rows[0]["WEBSITE_EI"] + "</font>                                                                                    \n");
            htmlStr.Append("				</td>                                                                                                                                                                                                    \n");
            htmlStr.Append("				<td class=xl6325524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl9725524>Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "</td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl9725524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl9725524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl9725524></td>                                                                                                                                                                                \n");
            htmlStr.Append("				<td class=xl7125524>&nbsp;</td>                                                                                                                                                                          \n");
            htmlStr.Append("			</tr>                                                                                                                                                                                                        \n");
            htmlStr.Append("		<tr height=22 style='mso-height-source: userset; height: 21.1875pt'>                                                                                                                                               \n");
            htmlStr.Append("			<td colspan=11 height=22 class=xl11325524                                                                                                                                                                    \n");
            htmlStr.Append("				style='border-right: 1.0pt solid black; height: 21.1875pt'>(C&#7847;n                                                                                                                                      \n");
            htmlStr.Append("				ki&#7875;m tra, &#273;&#7889;i chi&#7871;u khi l&#7853;p, giao,                                                                                                                                          \n");
            htmlStr.Append("				nh&#7853;n hóa &#273;&#417;n)</td>                                                                                                                                                                       \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<tr height=21 style='height: 19.5pt'>                                                                                                                                                                            \n");
            htmlStr.Append("			<td colspan=11 height=21 class=xl11625524 style='height: 19.5pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                                                                     \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<![if supportMisalignedColumns]>                                                                                                                                                                                 \n");
            htmlStr.Append("		<tr height=0 style='display: none'>                                                                                                                                                                              \n");
            htmlStr.Append("			<td width=38 style='width: 35pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=81 style='width: 76.25pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=84 style='width: 78.75pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=62 style='width: 58.75pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=84 style='width: 78.75pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=41 style='width: 38.75pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=74 style='width: 68.75pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=106 style='width: 100pt'></td>                                                                                                                                                                      \n");
            htmlStr.Append("			<td width=52 style='width: 48.75pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=62 style='width: 58.75pt'></td>                                                                                                                                                                       \n");
            htmlStr.Append("			<td width=12 style='width: 11.25pt'></td>                                                                                                                                                                        \n");
            htmlStr.Append("		</tr>                                                                                                                                                                                                            \n");
            htmlStr.Append("		<![endif]>                                                                                                                                                                                                       \n");
            htmlStr.Append("	</table>                                                                                                                                                                                                             \n");
            htmlStr.Append("</body>                                                                                                                                                                                                 \n");
            htmlStr.Append("</html>               \n");

            string filePath = "";

            /*using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\webproject\e-invoice-ws\02.Web\AttachFileText\" + tei_einvoice_m_pk + ".html"))
            {
                file.WriteLine(htmlStr.ToString()); // "sb" is the StringBuilder
            }*/

            connection.Close();
            connection.Dispose();
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