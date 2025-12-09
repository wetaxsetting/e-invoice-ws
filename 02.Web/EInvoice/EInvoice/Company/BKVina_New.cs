
using iTextSharp.text.pdf;
using System;
using System.Data;
//using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;

using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
using System.Drawing.Text;
using System.Drawing.Imaging;
///using Oracle.ManagedDataAccess.Client;
//using Oracle.ManagedDataAccess.Types;

namespace EInvoice.Company
{
    public class BKVina_New
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
            int[] page = new int[50] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

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
            // string read_prive = "", read_en = "", read_amount = "", amout_vat = "";
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


            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																						 \n");
            htmlStr.Append("<html>                                                                                                                                                                                       \n");
            htmlStr.Append("<head>                                                                                                                                                                                       \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                                          \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append("<script type='text/javascript' src='E:/Tomcat/webapps/e-invoice/system/syscommand.js'></script>                                                                                       \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                                              \n");
            htmlStr.Append(" <!-- Normalize or reset CSS with your favorite library -->                                                                                                                                  \n");
            htmlStr.Append("  <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'>                                                                                        \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            //htmlStr.Append("  <!-- Load paper.css for happy printing -->                                                                                                                                                 \n");
            //htmlStr.Append("  <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>                                                                                            \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append("  <!-- Set page size here: A5, A4 or A3 -->                                                                                                                                                  \n");
            htmlStr.Append("  <!-- Set also 'landscape' if you need -->                                                                                                                                                  \n");
            htmlStr.Append("  <style>@page { size: A4; background: white;margin-left: 50mm  }</style>                                                                                                                                                          \n");
            htmlStr.Append("  <link href='https://fonts.googleapis.com/css?family=Tangerine:700' rel='stylesheet' type='text/css'>                                                                                       \n");
            htmlStr.Append("  <style>                                                                                                                                                                                    \n");
            htmlStr.Append("    /*body   { font-family: serif }                                                                                                                                                          \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                                          \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                                          \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                                             \n");
            htmlStr.Append("    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                                                                             \n");
            htmlStr.Append("    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                                                                             \n");
            htmlStr.Append("    li     { font-size: 11pt; line-height: 5mm }                                                                                                                                             \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                                    \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                                            \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                                             \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                                  \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                                         \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                                        \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                                         \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                                            \n");
            htmlStr.Append("    body {                                                                                                                                                                                   \n");
            htmlStr.Append("       		 color: blue;                                                                                                                                                                    \n");
            htmlStr.Append("       		 font-size:100%;                                                                                                                                                                 \n");
            htmlStr.Append("       		 background-image: white;                                                                                                                                   \n");
            htmlStr.Append("		 }                                                                                                                                                                                   \n");
            htmlStr.Append("	h1 {                                                                                                                                                                                     \n");
            htmlStr.Append("	        color: #00FF00;                                                                                                                                                                  \n");
            htmlStr.Append("	}                                                                                                                                                                                        \n");
            htmlStr.Append("	p {                                                                                                                                                                                      \n");
            htmlStr.Append("	        color: rgb(0,0,255)                                                                                                                                                              \n");
            htmlStr.Append("	}                                                                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                                                                         \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append("   <!--table                                                                                                                                                                                 \n");
            htmlStr.Append("	{mso-displayed-decimal-separator:'\\.';                                                                                                                                                  \n");
            htmlStr.Append("	mso-displayed-thousand-separator:'\\,';}                                                                                                                                                 \n");
            htmlStr.Append("@page                                                                                                                                                                                        \n");
            htmlStr.Append("	{margin:2in 0in 0in 0in;                                                                                                                                                                 \n");
            htmlStr.Append("	mso-header-margin:0in;                                                                                                                                                                   \n");
            htmlStr.Append("	mso-footer-margin:0in;                                                                                                                                                                   \n");
            htmlStr.Append("	mso-horizontal-page-align:center;                                                                                                                                                        \n");
            htmlStr.Append("	size: A4; background: white;                                                                                                                                                             \n");
            htmlStr.Append("	mso-vertical-page-align:center;}                                                                                                                                                         \n");
            htmlStr.Append("tr                                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-height-source:auto;}                                                                                                                                                                \n");
            htmlStr.Append("col                                                                                                                                                                                          \n");
            htmlStr.Append("	{mso-width-source:auto;}                                                                                                                                                                 \n");
            htmlStr.Append("br                                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-data-placement:same-cell;}                                                                                                                                                          \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append(" @font-face {                                                                                                                                                                                \n");
            htmlStr.Append("    font-family:free3of9;                                                                                                                                                                    \n");
            htmlStr.Append("    src: url('.../e-invoice-ws/font/free3of9.woff');                                                                                                                                               \n");
            htmlStr.Append("	                                                                                                                                                                                         \n");
            htmlStr.Append("	}	                                                                                                                                                                                     \n");
            htmlStr.Append("	                                                                                                                                                                                         \n");
            htmlStr.Append(".style0                                                                                                                                                                                      \n");
            htmlStr.Append("	{mso-number-format:General;                                                                                                                                                              \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-rotate:0;                                                                                                                                                                            \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                              \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                        \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                        \n");
            htmlStr.Append("	font-size:18.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                       \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                    \n");
            htmlStr.Append("	font-family:Arial;                                                                                                                                                                       \n");
            htmlStr.Append("	mso-generic-font-family:auto;                                                                                                                                                            \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border:none;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-protection:locked visible;                                                                                                                                                           \n");
            htmlStr.Append("	mso-style-name:Normal;                                                                                                                                                                   \n");
            htmlStr.Append("	mso-style-id:0;}                                                                                                                                                                         \n");
            htmlStr.Append(".font5                                                                                                                                                                                       \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                       \n");
            htmlStr.Append("	font-size:18.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                       \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                    \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append(".font8                                                                                                                                                                                       \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                       \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                    \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append(".font10                                                                                                                                                                                      \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                       \n");
            htmlStr.Append("	font-size:18.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                    \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append(".font15                                                                                                                                                                                      \n");
            htmlStr.Append("	{color:windowtext;                                                                                                                                                                       \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                       \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                    \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append(".font18                                                                                                                                                                                      \n");
            htmlStr.Append("	{color:red;                                                                                                                                                                              \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                       \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                    \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append("td                                                                                                                                                                                           \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	padding:0px;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-ignore:padding;                                                                                                                                                                      \n");
            htmlStr.Append("	color:windowtext;                                                                                                                                                                        \n");
            htmlStr.Append("	font-size:18.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:400;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:normal;                                                                                                                                                                       \n");
            htmlStr.Append("	text-decoration:none;                                                                                                                                                                    \n");
            htmlStr.Append("	font-family:Arial;                                                                                                                                                                       \n");
            htmlStr.Append("	mso-generic-font-family:auto;                                                                                                                                                            \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:General;                                                                                                                                                               \n");
            htmlStr.Append("	text-align:general;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:bottom;                                                                                                                                                                   \n");
            htmlStr.Append("	border:none;                                                                                                                                                                             \n");
            htmlStr.Append("	mso-background-source:auto;                                                                                                                                                              \n");
            htmlStr.Append("	mso-pattern:auto;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-protection:locked visible;                                                                                                                                                           \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-rotate:0;}                                                                                                                                                                           \n");
            htmlStr.Append(".xl65                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl66                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl67                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                                  \n");
            htmlStr.Append(".xl68                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:maroon;                                                                                                                                                                            \n");
            htmlStr.Append("	font-size:20.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl69                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:2.0pt double windowtext;                                                                                                                                                   \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl70                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl71                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:2.0pt double windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl72                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl73                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;}                                                                                                                                                                      \n");
            htmlStr.Append(".xl74                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl75                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-rotate:90;}                                                                                                                                                                          \n");
            htmlStr.Append(".xl76                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl77                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl78                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-rotate:90;}                                                                                                                                                                          \n");
            htmlStr.Append(".xl79                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl80                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl81                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl82                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl83                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl84                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl85                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl86                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl87                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl88                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl89                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl90                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl91                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl92                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl93                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl94                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl95                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl96                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl97                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl98                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl99                                                                                                                                                                                        \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl100                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl101                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl102                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl103                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl104                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl105                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl106                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;}                                                                                                                                                                      \n");
            htmlStr.Append(".xl107                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:#C00000;                                                                                                                                                                           \n");
            htmlStr.Append("	font-size:40.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;}                                                                                                                                                                  \n");
            htmlStr.Append(".xl108                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl109                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("  white-space:nowrap; \n");
             htmlStr.Append("	text-align:left;}                                                                                                                                                                        \n");
            htmlStr.Append(".xl110                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");

            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl111                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:2.0pt double windowtext;                                                                                                                                                   \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl112                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl113                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl114                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl115                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid black;                                                                                                                                                            \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl116                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl117                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl118                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl119                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	font-size:20pt;                                                                                                                                                                        \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid black;}                                                                                                                                                           \n");
            htmlStr.Append(".xl120                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl121                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl122                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl123                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl124                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl125                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl126                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl127                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl128                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:0%;                                                                                                                                                                    \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl129                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:0%;                                                                                                                                                                    \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl130                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl131                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid black;                                                                                                                                                          \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl132                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                                                                       \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl133                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:white;                                                                                                                                                                             \n");
            htmlStr.Append("	font-size:18.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	background:#C00000;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern:black none;}                                                                                                                                                                 \n");
            htmlStr.Append(".xl134                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:white;                                                                                                                                                                             \n");
            htmlStr.Append("	font-size:14.0pt;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	background:#C00000;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern:black none;}                                                                                                                                                                 \n");
            htmlStr.Append(".xl135                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl136                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl137                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl138                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl139                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl140                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl141                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid black;}                                                                                                                                                           \n");
            htmlStr.Append(".xl142                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl143                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid black;}                                                                                                                                                           \n");
            htmlStr.Append(".xl144                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl145                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl146                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid black;}                                                                                                                                                           \n");
            htmlStr.Append(".xl147                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid windowtext;                                                                                                                                                        \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl148                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl149                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl150                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid black;}                                                                                                                                                           \n");
            htmlStr.Append(".xl151                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl152                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid black;}                                                                                                                                                           \n");
            htmlStr.Append(".xl153                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl154                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl155                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl156                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl157                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                                       \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl158                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid black;                                                                                                                                                           \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl159                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl160                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl161                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid black;}                                                                                                                                                           \n");
            htmlStr.Append(".xl162                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl163                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl164                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl165                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-rotate:90;}                                                                                                                                                                          \n");
            htmlStr.Append(".xl166                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl167                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;}                                                                                                                                                      \n");
            htmlStr.Append(".xl168                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;                                                                                                                                                                        \n");
            htmlStr.Append("	mso-rotate:90;}                                                                                                                                                                          \n");
            htmlStr.Append(".xl169                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:30.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	                                                                                                                                                                                         \n");
            htmlStr.Append("	                                                                                                                                                                                         \n");
            htmlStr.Append("	}                                                                                                                                                                                        \n");
            htmlStr.Append(".xl170                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:31.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;}                                                                                                                                                                      \n");
            htmlStr.Append(".xl171                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:27.0pt;                                                                                                                                                                        \n");
            htmlStr.Append("	font-weight:700;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;}                                                                                                                                                                      \n");
            htmlStr.Append(".xl172                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl173                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl174                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	white-space:normal;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl175                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0';                                                                                                                                                       \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl176                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:right;                                                                                                                                                                        \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl177                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	white-space:nowrap;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-text-control:shrinktofit;}                                                                                                                                                           \n");
            htmlStr.Append(".xl178                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl179                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl180                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'_\\(* \\#\\,\\#\\#0_\\)\\;_\\(* \\\\\\(\\#\\,\\#\\#0\\\\\\)\\;_\\(* \\0022-\\0022??_\\)\\;_\\(\\@_\\)';                                                               \n");
            htmlStr.Append("	text-align:center;}                                                                                                                                                                      \n");
            htmlStr.Append(".xl181                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	mso-number-format:'\\#\\,\\#\\#0_\\)\\;\\\\\\(\\#\\,\\#\\#0\\\\\\)';                                                                                                                     \n");
            htmlStr.Append("	text-align:right;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl182                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	white-space:normal;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl183                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl184                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl185                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	color:red;                                                                                                                                                                               \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;}                                                                                                                                                                     \n");
            htmlStr.Append(".xl186                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:12pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:center;                                                                                                                                                                       \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:1.0pt solid windowtext;                                                                                                                                                     \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl187                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	border-top:2.0pt double windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl188                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	text-align:left;                                                                                                                                                                         \n");
            htmlStr.Append("	border-top:2.0pt double windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl189                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:top;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl190                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	border-top:1.0pt solid black;                                                                                                                                                             \n");
            htmlStr.Append("	border-right:none;                                                                                                                                                                       \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:none;}                                                                                                                                                                       \n");
            htmlStr.Append(".xl191                                                                                                                                                                                       \n");
            htmlStr.Append("	{mso-style-parent:style0;                                                                                                                                                                \n");
            htmlStr.Append("	font-size:16.6pt;                                                                                                                                                                         \n");
            htmlStr.Append("	font-style:italic;                                                                                                                                                                       \n");
            htmlStr.Append("	font-family:Arial, sans-serif;                                                                                                                                                           \n");
            htmlStr.Append("	mso-font-charset:0;                                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align:middle;                                                                                                                                                                   \n");
            htmlStr.Append("	border-top:none;                                                                                                                                                                         \n");
            htmlStr.Append("	border-right:1.0pt solid windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	border-bottom:none;                                                                                                                                                                      \n");
            htmlStr.Append("	border-left:1.0pt solid windowtext;                                                                                                                                                       \n");
            htmlStr.Append("	mso-rotate:90;}                                                                                                                                                                          \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append("	                                                                                                                                                                                         \n");
            htmlStr.Append("#rotate {                                                                                                                                                                                    \n");
            htmlStr.Append("     -moz-transform: rotate(-90.0deg);  /* FF3.5+ */                                                                                                                                         \n");
            htmlStr.Append("       -o-transform: rotate(-90.0deg);  /* Opera 10.5 */                                                                                                                                     \n");
            htmlStr.Append("  -webkit-transform: rotate(-90.0deg);  /* Saf3.1+, Chrome */                                                                                                                                \n");
            htmlStr.Append("             filter:  progid:DXImageTransform.Microsoft.BasicImage(rotation=0.083);  /* IE6,IE7 */                                                                                           \n");
            htmlStr.Append("         -ms-filter: 'progid:DXImageTransform.Microsoft.BasicImage(rotation=0.083)'; /* IE8 */                                                                                               \n");
            htmlStr.Append("}                                                                                                                                                                                            \n");
            htmlStr.Append("-->                                                                                                                                                                                          \n");
            htmlStr.Append("</style>                                                                                                                                                                                     \n");
            htmlStr.Append("                                                                                                                                                                                             \n");
            htmlStr.Append("</head>                                                                                                                                                                                      \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                                            \n");
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

            double v_totalHeightLastPage = 550;// 258.5;

            double v_totalHeightPage = 580;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 590;// 540;

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
                htmlStr.Append("	                                                                                                                                                      \n");
                htmlStr.Append("		<table border=0 cellpadding=0 cellspacing=0 width=733 style='border-collapse:                                                                                                        \n");
                htmlStr.Append("		 collapse;table-layout:fixed;width:1060.32pt'>                                                                                                                                           \n");
                htmlStr.Append("		 <col class=xl65 width=3 style='mso-width-source:userset;mso-width-alt:113;                                                                                                          \n");
                htmlStr.Append("		 width:3.36pt'>                                                                                                                                                                         \n");
                htmlStr.Append("		 <col class=xl65 width=5 style='mso-width-source:userset;mso-width-alt:170;                                                                                                          \n");
                htmlStr.Append("		 width:6.72pt'>                                                                                                                                                                         \n");
                htmlStr.Append("		 <col class=xl65 width=29 style='mso-width-source:userset;mso-width-alt:1024;                                                                                                        \n");
                htmlStr.Append("		 width:36.96pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=86 style='mso-width-source:userset;mso-width-alt:3072;                                                                                                        \n");
                htmlStr.Append("		 width:109.2pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=66 style='mso-width-source:userset;mso-width-alt:2332;                                                                                                        \n");
                htmlStr.Append("		 width:82.32pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=45 style='mso-width-source:userset;mso-width-alt:1592;                                                                                                        \n");
                htmlStr.Append("		 width:47.6pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=104 style='mso-width-source:userset;mso-width-alt:3697;                                                                                                       \n");
                htmlStr.Append("		 width:109.2pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=22 style='mso-width-source:userset;mso-width-alt:796;                                                                                                         \n");
                htmlStr.Append("		 width:28.56pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=51 style='mso-width-source:userset;mso-width-alt:1820;                                                                                                        \n");
                htmlStr.Append("		 width:63.84pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=14 style='mso-width-source:userset;mso-width-alt:483;                                                                                                         \n");
                htmlStr.Append("		 width:16.8pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=17 style='mso-width-source:userset;mso-width-alt:597;                                                                                                         \n");
                htmlStr.Append("		 width:21.84pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=36 style='mso-width-source:userset;mso-width-alt:1280;                                                                                                        \n");
                htmlStr.Append("		 width:45.36pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=35 style='mso-width-source:userset;mso-width-alt:1251;                                                                                                        \n");
                htmlStr.Append("		 width:43.68pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=20 style='mso-width-source:userset;mso-width-alt:711;                                                                                                         \n");
                htmlStr.Append("		 width:25.2pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=35 style='mso-width-source:userset;mso-width-alt:1251;                                                                                                        \n");
                htmlStr.Append("		 width:43.68pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=22 style='mso-width-source:userset;mso-width-alt:796;                                                                                                         \n");
                htmlStr.Append("		 width:28.56pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=53 style='mso-width-source:userset;mso-width-alt:1877;                                                                                                        \n");
                htmlStr.Append("		 width:67.2pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl65 width=68 style='mso-width-source:userset;mso-width-alt:2417;                                                                                                        \n");
                htmlStr.Append("		 width:85.68t'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <col class=xl66 width=22 style='mso-width-source:userset;mso-width-alt:796;                                                                                                         \n");
                htmlStr.Append("		 width:35.56pt'>                                                                                                                                                                        \n");
                htmlStr.Append("		 <tr height=40 style='mso-outline-parent:collapsed;height:36.0pt'>                                                                                                                   \n");
                htmlStr.Append("		  <td height=40 class=xl65 width=3 style='height:36.0pt;width:2.65pt'></td>                                                                                                             \n");
                htmlStr.Append("		  <td class=xl65 width=5 style='width:5.3pt'></td>                                                                                                                                     \n");
                htmlStr.Append("		  <td width=29 style='width:28.00pt' align=left valign=top><span style='mso-ignore:vglayout;                                                                                            \n");
                htmlStr.Append("		  position:absolute;z-index:2;margin-left:26px;margin-top:3px;width:250px;                                                                                                           \n");
                htmlStr.Append("		  height:180px'><img width=250 height=180 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/image007.png' v:shapes='Picture_x0020_3'></span><![endif]><span                      \n");
                htmlStr.Append("		  style='mso-ignore:vglayout2'>                                                                                                                                                      \n");
                htmlStr.Append("		  <table cellpadding=0 cellspacing=0>                                                                                                                                                \n");
                htmlStr.Append("		   <tr>                                                                                                                                                                              \n");
                htmlStr.Append("		    <td height=40 class=xl65 width=29 style='height:36.0pt;width:28.00pt'></td>                                                                                                         \n");
                htmlStr.Append("		   </tr>                                                                                                                                                                             \n");
                htmlStr.Append("		  </table>                                                                                                                                                                           \n");
                htmlStr.Append("		  <td class=xl65 width=86 style='width:86.00pt'></td>                                                                                                                                   \n");
                htmlStr.Append("		  </span></td>                                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=13 class=xl107 width=499 style='width:375pt'>" + dt.Rows[0]["Seller_Name"] + " </td>                                                                                                       \n");
           
                htmlStr.Append("		  <td class=xl65 width=66 style='width:64.84pt'></td>                                                                                                                                   \n");
               
                htmlStr.Append("		  <td class=xl65 width=45 style='width:37.54pt'></td>                                                                                                                                   \n");


                htmlStr.Append("		 <tr height=16 style='mso-height-source:userset;height:14.4pt'>                                                                                                                      \n");
                htmlStr.Append("		  <td height=16 class=xl65 style='height:14.4pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65 colspan=2></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65 colspan=1>&#272;&#7883;a ch&#7881; <font class='font8'>(Address):</font></td>                                                                                                 \n");

                htmlStr.Append("		  <td class=xl65 colspan=1></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=11 rowspan=2 class=xl108 width=395 style='width:297pt'>" + dt.Rows[0]["Seller_Address"] + " </td>                                                                                        \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");

                htmlStr.Append("		 <tr height=20 style='mso-height-source:userset;height:18.0pt'>                                                                                                                      \n");
                htmlStr.Append("		  <td height=20 class=xl65 style='height:18.0pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl68></td>                                                                                                                                                               \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");

                htmlStr.Append("		 <tr height=16 style='mso-height-source:userset;height:14.4pt'>                                                                                                                      \n");
                htmlStr.Append("		  <td height=16 class=xl65 style='height:14.4pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65  colspan=2></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65 colspan=2 style='mso-ignore:colspan'>S&#7889; tài kho&#7843;n <font                                                                                                 \n");
                htmlStr.Append("		  class='font8'>(Bank Account):</font></td>                                                                                                                                          \n");
                htmlStr.Append("		  <td class=xl65 colspan=1></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=10 class=xl109>" + dt.Rows[0]["Seller_AccountNo"] + " - " + dt.Rows[0]["Seller_AccountNo2"] + "</td>                                                                    \n");

                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");

                htmlStr.Append("		 <tr height=16 style='mso-height-source:userset;height:14.4pt'>                                                                                                                      \n");
                htmlStr.Append("		  <td height=16 class=xl65 style='height:14.4pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65 colspan=2></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65 colspan=1 style='mso-ignore:colspan'>T&#7841;i <font                                                                                                 \n");
                htmlStr.Append("		  class='font8'>(At):</font></td>                                                                                                                                          \n");
                htmlStr.Append("		  <td class=xl65 colspan=1></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=8 class=xl109> " + dt.Rows[0]["Seller_Bankname"] + "</td>                                                                    \n");


                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");

                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr height=16 style='mso-height-source:userset;height:14.4pt'>                                                                                                                      \n");
                htmlStr.Append("		  <td height=16 class=xl65 style='height:14.4pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65 colspan=2></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=12 class=xl110>&#272;i&#7879;n tho&#7841;i <font class='font8'>(Tel):                                                                                                  \n");
                htmlStr.Append("		  </font><font class='font5'>0274-3559-826 ~ 30</font><span                                                                                                                          \n");
                htmlStr.Append("		  style='mso-spacerun:yes'>   </span>Mã s&#7889; thu&#7871; <font class='font8'>(Tax Code):</font>                                                                                   \n");
                htmlStr.Append("		  " + dt.Rows[0]["Seller_TaxCode"] + "</td>                                                                                                                                                                    \n");

                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");

                htmlStr.Append("		 <tr height=2 style='mso-height-source:userset;height:2.25pt'>                                                                                                                       \n");
                htmlStr.Append("		  <td height=2 class=xl65 style='height:2.25pt'></td>                                                                                                                                \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl69>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl69>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl69>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl69>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td colspan=12 class=xl111>&nbsp;</td>                                                                                                                                             \n");
                htmlStr.Append("		  <td class=xl70></td>                                                                                                                                                               \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr height=42 style='mso-height-source:userset;height:31.2pt'>                                                                                                                      \n");
                htmlStr.Append("		  <td height=42 class=xl65 style='height:31.2pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl71>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td align=left valign=top colspan=3><span style='mso-ignore:vglayout;                                                                                                             \n");
                htmlStr.Append("		    position: absolute; z-index:1; margin-left:0px; margin-top:-10px; width: 446px;                                                                                           \n");
                htmlStr.Append("		          height: 337px'><img width=275.0 height=25 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/Barcode_Code.png' v:shapes='Rectangle_x0020_1'></span><![endif]><span                        \n");
                htmlStr.Append("		    style = 'mso-ignore:vglayout2'>                                                                                                                                                \n");
                htmlStr.Append("		    <table cellpadding = 0 cellspacing = 0>                                                                                                                                       \n");
                htmlStr.Append("		        <tr>                                                                                                                                                                      \n");
                htmlStr.Append("		         <td height = 18 class=xl97 width = 66 style='height:15.84‬pt;width:64.84pt'></td>                                                                                           \n");
                htmlStr.Append("		     </tr>                                                                                                                                                                          \n");
                htmlStr.Append("		    </table>                                                                                                                                                                        \n");
                htmlStr.Append("		    </span></td>                                                                        																							\n");





                htmlStr.Append("		  <td class=xl169></td>                                                                                                                                                              \n");
                htmlStr.Append("		  <td colspan=6 class=xl170>HOÁ &#272;&#416;N<span                                                                                                                                   \n");
                htmlStr.Append("		  style='mso-spacerun:yes'> </span></td>                                                                                                                                             \n");
                htmlStr.Append("		  <td class=xl171></td>                                                                                                                                                              \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=3 class=xl109><font class='font8'></font></td>                                                                                            \n");
                htmlStr.Append("		  <td class=xl187 colspan=2 style='mso-ignore:colspan;border-right:1.0pt solid black'></td>                                                                                \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr height=32 style='height:24.0pt'>                                                                                                                                                \n");
                htmlStr.Append("		  <td height=32 class=xl65 style='height:24.0pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl72>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=6 class=xl170>GIÁ TR&#7882; GIA T&#258;NG</td>                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl109 colspan=3 style='mso-ignore:colspan'>Ký hi&#7879;u <font                                                                                                           \n");
                htmlStr.Append("		  class='font8'>(Serial):</font></td>                                                                                                                                                \n");
                htmlStr.Append("		  <td class=xl65>" + dt.Rows[0]["templateCode"] + " " + dt.Rows[0]["InvoiceSerialNo"] + " </td>                                                                                                                                                  \n");
                htmlStr.Append("		  <td class=xl74>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr height=26 style='mso-height-source:userset;height:23.4pt'>                                                                                                                      \n");
                htmlStr.Append("		  <td height=26 class=xl65 style='height:23.4pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl72>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=7 class=xl109></td>                                                                                                                                                    \n");
                htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl109 colspan=3 style='mso-ignore:colspan'>S&#7889; <font                                                                                                                \n");
                htmlStr.Append("		  class='font8'>(Invoice no):</font></td>                                                                                                                                            \n");
                htmlStr.Append("		  <td class=xl109>" + dt.Rows[0]["InvoiceNumber"] + " </td>                                                                                                                                                \n");
                htmlStr.Append("		  <td class=xl74>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr class=xl97 height=30 style='mso-height-source:userset;height:27.54‬pt'>                                                                                                          \n");
                htmlStr.Append("		  <td height=30 class=xl97 style='height:27.54‬pt'></td>                                                                                                                              \n");
                htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td colspan=7 class=xl172>Ngày <font class='font8'>Date</font><font                                                                                                                \n");
                htmlStr.Append("		  class='font5'><span style='mso-spacerun:yes'>  " + dt.Rows[0]["invoiceissueddate_dd"] + "   </span>Tháng </font><font                                                                                               \n");
                htmlStr.Append("		  class='font8'>Month</font><font class='font5'><span                                                                                                                                \n");
                htmlStr.Append("		  style='mso-spacerun:yes'>  " + dt.Rows[0]["invoiceissueddate_mm"] + "   </span>N&#259;m </font><font class='font8'>Years</font><font                                                                                \n");
                htmlStr.Append("		  class='font5'><span style='mso-spacerun:yes'> " + dt.Rows[0]["invoiceissueddate_yyyy"] + " </span></font></td>                                                                                                      \n");
                htmlStr.Append("		                                                                                                                                                                                     \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl166>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr class=xl97 height=18 style='height:15.84‬pt'>                                                                                                                                    \n");
                htmlStr.Append("		  <td height=18 class=xl97 style='height:15.84‬pt'></td>                                                                                                                              \n");
                htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		  <td class=xl97 colspan=8 style='mso-ignore:colspan'>H&#7885; tên                                                                                                                   \n");
                htmlStr.Append("		  ng&#432;&#7901;i mua hàng <font class='font8'>(Customer's name</font><font                                                                                                         \n");
                htmlStr.Append("		  class='font10'>):</font><font class='font5'> " + dt.Rows[0]["buyer"] + "</font></td>                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                if (dt.Rows[0]["BuyerLegalName"].ToString().Length >= 70)
                {
                    htmlStr.Append("		 <tr class=xl97 height=18 style='mso-height-source:userset;height:15.84‬pt'>                                                                                                           \n");
                    htmlStr.Append("		  <td height=18 class=xl97 style='height:15.84‬pt'></td>                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		  <td class=xl97 colspan=3 style='mso-ignore:colspan'>Tên &#273;&#417;n                                                                                                              \n");
                    htmlStr.Append("		  v&#7883;<font class='font10'> </font><font class='font8'>(Company's name)</font><font                                                                                              \n");
                    htmlStr.Append("		  class='font15'>:</font></td>                                                                                                                                                       \n");
                    htmlStr.Append("		  <td colspan=13 rowspan=2 class=xl108 width=522 style='width:392pt'>" + dt.Rows[0]["buyerlegalname"] + " </td>                                                                      \n");
                    htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                    htmlStr.Append("		 <tr class=xl97 height=12 style='mso-height-source:userset;height:10.8pt'>                                                                                                            \n");
                    htmlStr.Append("		  <td height=12 class=xl97 style='height:10.8pt'></td>                                                                                                                                \n");
                    htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                }
                else
                {
                    htmlStr.Append("		 <tr class=xl97 height=18 style='mso-height-source:userset;height:15.84‬pt'>                                                                                                           \n");
                    htmlStr.Append("		  <td height=18 class=xl97 style='height:15.84‬pt'></td>                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		  <td class=xl97 colspan=3 style='mso-ignore:colspan'>Tên &#273;&#417;n                                                                                                              \n");
                    htmlStr.Append("		  v&#7883;<font class='font10'> </font><font class='font8'>(Company's name)</font><font                                                                                              \n");
                    htmlStr.Append("		  class='font15'>:</font></td>                                                                                                                                                       \n");
                    htmlStr.Append("		  <td colspan=13 class=xl108 width=522 style='width:392pt'>" + dt.Rows[0]["buyerlegalname"] + " </td>                                                                                \n");
                    htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                }
                if (dt.Rows[0]["buyeraddress"].ToString().Length > 90)
                {
                    htmlStr.Append("		 <tr class=xl97 height=18 style='mso-height-source:userset;height:15.84‬pt'>                                                                                                           \n");
                    htmlStr.Append("		  <td height=18 class=xl97 style='height:15.84‬pt'></td>                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		  <td class=xl97 colspan=2 style='mso-ignore:colspan'>&#272;&#7883;a ch&#7881; <font                                                                                                 \n");
                    htmlStr.Append("		  class='font8'>(Address):</font></td>                                                                                                                                               \n");
                    htmlStr.Append("		  <td colspan=14 rowspan=2 class=xl108 width=588 style='width:441pt'>" + dt.Rows[0]["buyeraddress"] + " </td>                                                                        \n");
                    htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                    htmlStr.Append("		 <tr class=xl97 height=12 style='mso-height-source:userset;height:10.8pt'>                                                                                                            \n");
                    htmlStr.Append("		  <td height=12 class=xl97 style='height:10.8pt'></td>                                                                                                                                \n");
                    htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl97></td>                                                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                }
                else
                {
                    htmlStr.Append("		 <tr class=xl97 height=18 style='mso-height-source:userset;height:15.84‬pt'>                                                                                                           \n");
                    htmlStr.Append("		  <td height=18 class=xl97 style='height:15.84‬pt'></td>                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		  <td class=xl97 colspan=2 style='mso-ignore:colspan'>&#272;&#7883;a ch&#7881; <font                                                                                                 \n");
                    htmlStr.Append("		  class='font8'>(Address):</font></td>                                                                                                                                               \n");
                    htmlStr.Append("		  <td colspan=14 class=xl108 width=588 style='width:441pt'>" + dt.Rows[0]["buyeraddress"] + " </td>                                                                                  \n");
                    htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                }
                htmlStr.Append("		 <tr class=xl97 height=18 style='height:15.84‬pt'>                                                                                                                                     \n");
                htmlStr.Append("		  <td height=18 class=xl97 style='height:15.84‬pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		  <td class=xl97 colspan=2 style='mso-ignore:colspan'>Mã s&#7889; thu&#7871; <font                                                                                                   \n");
                htmlStr.Append("		  class='font8'>(Tax co</font><font class='font8'>de)</font><font                                                                                                                    \n");
                htmlStr.Append("		  class='font15'>:</font></td>                                                                                                                                                       \n");
                htmlStr.Append("		  <td align=left valign=top><span style='mso-ignore:vglayout;                                                                                                                        \n");
                htmlStr.Append("		  position:absolute;z-index:1;margin-left:0px;margin-top:-15px;width:860px;                                                                                                          \n");
                htmlStr.Append("		  height:750px'><img width=860 height=750 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/image008.png' v:shapes='Rectangle_x0020_1'></span><![endif]><span                    \n");
                htmlStr.Append("		  style='mso-ignore:vglayout2'>                                                                                                                                                      \n");
                htmlStr.Append("		  <table cellpadding=0 cellspacing=0>                                                                                                                                                \n");
                htmlStr.Append("		   <tr>                                                                                                                                                                              \n");
                htmlStr.Append("		    <td height=18 class=xl97 width=66 style='height:15.84‬pt;width:64.84pt'></td>                                                                                                      \n");
                htmlStr.Append("		   </tr>                                                                                                                                                                             \n");
                htmlStr.Append("		  </table>                                                                                                                                                                           \n");
                htmlStr.Append("		  </span></td>                                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=13 class=xl173>" + dt.Rows[0]["BuyerTaxCode"] + " </td>                                                                                                                                       \n");
                htmlStr.Append("		  <td class=xl165>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr class=xl97 height=26 style='mso-height-source:userset;height:23.94‬pt'>                                                                                                          \n");
                htmlStr.Append("		  <td height=26 class=xl97 style='height:23.94‬pt'></td>                                                                                                                              \n");
                htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		  <td class=xl97 colspan=5 style='mso-ignore:colspan'>Hình th&#7913;c thanh                                                                                                          \n");
                htmlStr.Append("		  toán (<font class='font8'>Mode of payment</font><font class='font10'>):</font><font                                                                                                \n");
                htmlStr.Append("		  class='font5'><span style='mso-spacerun:yes'>  " + dt.Rows[0]["PaymentMethodCK"] + " </span></font></td>                                                                           \n");
                htmlStr.Append("		  <td colspan=2 class=xl173></td>                                                                                                                                                    \n");
                htmlStr.Append("		  <td class=xl97 colspan=7 style='mso-ignore:colspan'>S&#7889; tài kho&#7843;n<font                                                                                                  \n");
                htmlStr.Append("		  class='font10'> </font><font class='font8'>(Bank Account)</font><font                                                                                                              \n");
                htmlStr.Append("		  class='font15'>:</font></td>                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=2 class=xl173>" + dt.Rows[0]["BuyerAccountNo"] + " </td>                                                                                                               \n");
                htmlStr.Append("		  <td class=xl189 align=left valign=top><span style='mso-ignore:vglayout;                                                                                                            \n");
                htmlStr.Append("		  position:absolute;z-index:3;margin-left:4px;margin-top:25px;width:50px;                                                                                                            \n");
                htmlStr.Append("		  height:1250‬px‬'><img width=50 height=1250 src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/image004.png' v:shapes='Picture_x0020_378'></span><![endif]><span                            \n");
                htmlStr.Append("		  style='mso-ignore:vglayout2'>                                                                                                                                                      \n");
                htmlStr.Append("		                                                                                                                                                                                     \n");
                htmlStr.Append("		  </span></td>                                                                                                                                                                       \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
               
                htmlStr.Append("		 <tr class=xl97 height=26 style='mso-height-source:userset;height:23.94‬pt'>                                                                                                          \n");
                htmlStr.Append("		  <td height=26 class=xl97 style='height:23.94‬pt'></td>                                                                                                                              \n");
                htmlStr.Append("		  <td class=xl164>&nbsp;</td>                                                                                                                                                        \n");
                htmlStr.Append("		  <td colspan=5 class=xl173>Đơn vị tiền tệ<font                                                                                                  \n");
                htmlStr.Append("		  class='font10'> </font><font class='font8'>(Currency)</font><font                                                                                                              \n");
                htmlStr.Append("		  class='font15'>: " + dt.Rows[0]["CurrencyCodeUSD"] + "</font></td>                                                                                                               \n");
                htmlStr.Append("		  <td colspan=2 class=xl173></td>                                                                                                                                                    \n");
                htmlStr.Append("		  <td class=xl97 colspan=7 style='mso-ignore:colspan'><font                                                                                                  \n");
                htmlStr.Append("		  class='font10'> </font><font class='font8'></font><font                                                                                                              \n");
                htmlStr.Append("		  class='font15'></font></td>                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=2 class=xl173></td>                                                                                                               \n");
                htmlStr.Append("		  <td class=xl189 align=left valign=top></td>                                                                                                                                                                       \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                                \n");

               


                htmlStr.Append("		 <tr class=xl67 height=40 style='mso-height-source:userset;height:36.0pt'>                                                                                                           \n");
                htmlStr.Append("		  <td height=40 class=xl67 style='height:36.0pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl76>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl77 width=29 style='width:28.00pt'>STT<br>                                                                                                                               \n");
                htmlStr.Append("		    <font class='font8'>No.</font></td>                                                                                                                                              \n");
                htmlStr.Append("		  <td colspan=4 class=xl113 width=301 style='border-right:1.0pt solid black;                                                                                                          \n");
                htmlStr.Append("		  border-left:none;width:226pt'>Tên hàng hoá, d&#7883;ch v&#7909;<br>                                                                                                                \n");
                htmlStr.Append("		    <font class='font8'>Name of goods, services</font></td>                                                                                                                          \n");
                htmlStr.Append("		  <td colspan=2 class=xl115 width=73 style='border-right:1.0pt solid black;                                                                                                           \n");
                htmlStr.Append("		  border-left:none;width:55pt'>&#272;&#417;n v&#7883; tính<br>                                                                                                                       \n");
                htmlStr.Append("		    <font class='font8'>Unit</font></td>                                                                                                                                             \n");
                htmlStr.Append("		  <td colspan=3 class=xl115 width=67 style='border-right:1.0pt solid black;                                                                                                           \n");
                htmlStr.Append("		  border-left:none;width:50pt'>S&#7889; l&#432;&#7907;ng<br>                                                                                                                         \n");
                htmlStr.Append("		    <font class='font8'>Quantity</font></td>                                                                                                                                         \n");
                htmlStr.Append("		  <td colspan=3 class=xl115 width=90 style='border-right:1.0pt solid black;                                                                                                           \n");
                htmlStr.Append("		  border-left:none;width:67pt'>&#272;&#417;n giá<br>                                                                                                                                 \n");
                htmlStr.Append("		    <font class='font8'>Price</font></td>                                                                                                                                            \n");
                htmlStr.Append("		  <td colspan=3 class=xl115 width=143 style='border-right:1.0pt solid black;                                                                                                          \n");
                htmlStr.Append("		  border-left:none;width:108pt'>Thành ti&#7873;n<br>                                                                                                                                 \n");
                htmlStr.Append("		    <font class='font8'>Amount</font></td>                                                                                                                                           \n");
                htmlStr.Append("		  <td class=xl78>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                htmlStr.Append("		 <tr class=xl73 height=18 style='height:15.84‬pt'>                                                                                                                                     \n");
                htmlStr.Append("		  <td height=18 class=xl73 style='height:15.84‬pt'></td>                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl79>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		  <td class=xl80>1</td>                                                                                                                                                              \n");
                htmlStr.Append("		  <td colspan=4 class=xl117 style='border-right:1.0pt solid black;border-left:                                                                                                        \n");
                htmlStr.Append("		  none'>2</td>                                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=2 class=xl119 style='border-right:1.0pt solid black;border-left:                                                                                                        \n");
                htmlStr.Append("		  none'>3</td>                                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=3 class=xl119 style='border-right:1.0pt solid black;border-left:                                                                                                        \n");
                htmlStr.Append("		  none'>4</td>                                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=3 class=xl119 style='border-right:1.0pt solid black;border-left:                                                                                                        \n");
                htmlStr.Append("		  none'>5</td>                                                                                                                                                                       \n");
                htmlStr.Append("		  <td colspan=3 class=xl119 style='border-right:1.0pt solid black;border-left:                                                                                                        \n");
                htmlStr.Append("		  none'>6 = 4 x 5</td>                                                                                                                                                               \n");
                htmlStr.Append("		  <td class=xl78>&nbsp;</td>                                                                                                                                                         \n");
                htmlStr.Append("		 </tr>                                                                                                                                                                               \n");

                v_rowHeight = "50.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 23;

                v_rowHeightLast = "50.0pt";// "23.5pt";
                v_rowHeightLastNumber = 23;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "50.0pt"; //"26.5pt";    
                        v_rowHeightLast = "50.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 28;//27.5;
                        v_rowHeightEmptyLast = "22.0pt"; //"23.0pt";
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

                        htmlStr.Append("		 		<tr height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                              \n");
                        htmlStr.Append("				  <td height=21 class=xl65 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'></td>                                                                                                                      \n");
                        htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
                        htmlStr.Append("				  <td class=xl137 >" + dt_d.Rows[v_index][7] + "</td>                                                                                                                                            \n");
                        htmlStr.Append("				  <td colspan=4 class=xl138 width=301 style='border-right:1.0pt solid black;border-top:none;                                                                                  \n");
                        htmlStr.Append("				  border-left:none;width:226pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                                        \n");
                        htmlStr.Append("				  <td colspan=2 class=xl141 style='border-right:1.0pt solid black;border-left:none;border-top:none;'>" + dt_d.Rows[v_index][1] + "</td>                                                          \n");
                        htmlStr.Append("				  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;                                                                               \n");
                        htmlStr.Append("				  none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("				  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;                                                                               \n");
                        htmlStr.Append("				  none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("				  <td colspan=3 class=xl146 style='border-left:none;border-top:none;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                  \n");
                        htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("				</tr>                                                                                                                                                                        \n");
                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("				 		<tr height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                      \n");
                            htmlStr.Append("						  <td height=21 class=xl65 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'></td>                                                                                                             \n");
                            htmlStr.Append("						  <td class=xl81 style=' border-bottom:1.0pt solid black;'>&nbsp;</td>                                                                                               \n");
                            htmlStr.Append("						  <td class=xl137 style=' border-bottom:1.0pt solid black;' >" + dt_d.Rows[v_index][7] + "</td>                                                                                           \n");
                            htmlStr.Append("						  <td colspan=4 class=xl138 width=301 style='border-right:1.0pt solid black;border-top:none;border-bottom:1.0pt solid black;                                           \n");
                            htmlStr.Append("						  border-left:none;width:226pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                                \n");
                            htmlStr.Append("						  <td colspan=2 class=xl141 style='border-right:1.0pt solid black;border-left:none;border-top:none;border-bottom:1.0pt solid black;'>" + dt_d.Rows[v_index][1] + "</td>                   \n");
                            htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:1.0pt solid black;                                        \n");
                            htmlStr.Append("						  none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                        \n");
                            htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:1.0pt solid black;                                        \n");
                            htmlStr.Append("						  none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                                        \n");
                            htmlStr.Append("						  <td colspan=3 class=xl146 style='border-left:1.0pt solid black;;border-top:none;border-bottom:1.0pt solid black;border-right:1.0pt solid black;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>\n");
                            htmlStr.Append("						  <td class=xl191 style=' border-right:1.0pt solid black;'>&nbsp;</td>                                                                                                \n");
                            htmlStr.Append("						</tr>                                                                                                                                                                \n");

                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("				 		<tr height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                      \n");
                                htmlStr.Append("						  <td height=21 class=xl65 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-bottom:none;'></td>                                                                                                             \n");
                                htmlStr.Append("						  <td class=xl81 style=' border-bottom:none;'>&nbsp;</td>                                                                                               \n");
                                htmlStr.Append("						  <td class=xl137 style=' border-bottom:none;' >" + dt_d.Rows[v_index][7] + "</td>                                                                                           \n");
                                htmlStr.Append("						  <td colspan=4 class=xl138 width=301 style='border-right:1.0pt solid black;border-top:none;border-bottom:none;                                           \n");
                                htmlStr.Append("						  border-left:none;width:226pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                                \n");
                                htmlStr.Append("						  <td colspan=2 class=xl141 style='border-right:1.0pt solid black;border-left:none;border-top:none;border-bottom:none;'>" + dt_d.Rows[v_index][1] + "</td>                   \n");
                                htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:none;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                        \n");
                                htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:none;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                                        \n");
                                htmlStr.Append("						  <td colspan=3 class=xl146 style='border-left:1.0pt solid black;;border-top:none;border-bottom:none;border-right:1.0pt solid black;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>\n");
                                htmlStr.Append("						  <td class=xl191 style=' border-right:1.0pt solid black;'>&nbsp;</td>                                                                                                \n");
                                htmlStr.Append("						</tr>                                                                                                                                                                \n");

                            }
                            else
                            {
                                htmlStr.Append("				 		<tr height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                      \n");
                                htmlStr.Append("						  <td height=21 class=xl65 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";border-bottom:none;'></td>                                                                                                             \n");
                                htmlStr.Append("						  <td class=xl81 style=' border-bottom:none;'>&nbsp;</td>                                                                                               \n");
                                htmlStr.Append("						  <td class=xl137 style=' border-bottom:none;' >" + dt_d.Rows[v_index][7] + "</td>                                                                                           \n");
                                htmlStr.Append("						  <td colspan=4 class=xl138 width=301 style='border-right:1.0pt solid black;border-top:none;border-bottom:none;                                           \n");
                                htmlStr.Append("						  border-left:none;width:226pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                                \n");
                                htmlStr.Append("						  <td colspan=2 class=xl141 style='border-right:1.0pt solid black;border-left:none;border-top:none;border-bottom:none;'>" + dt_d.Rows[v_index][1] + "</td>                   \n");
                                htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:none;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                        \n");
                                htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:none;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                                        \n");
                                htmlStr.Append("						  <td colspan=3 class=xl146 style='border-left:1.0pt solid black;;border-top:none;border-bottom:none;border-right:1.0pt solid black;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>\n");
                                htmlStr.Append("						  <td class=xl191 style=' border-right:1.0pt solid black;'>&nbsp;</td>                                                                                                \n");
                                htmlStr.Append("						</tr>                                                                                                                                                                \n");

                            }

                        }
                    }
                    else
                    { // dong giua
                        htmlStr.Append("				<tr height=21 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                                                                              \n");
                        htmlStr.Append("				  <td height=21 class=xl65 style='height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'></td>                                                                                                                      \n");
                        htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
                        htmlStr.Append("				  <td class=xl137 >" + dt_d.Rows[v_index][7] + "</td>                                                                                                                                            \n");
                        htmlStr.Append("				  <td colspan=4 class=xl138 width=301 style='border-right:1.0pt solid black;border-top:none;                                                                                  \n");
                        htmlStr.Append("				  border-left:none;width:226pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                                        \n");
                        htmlStr.Append("				  <td colspan=2 class=xl141 style='border-right:1.0pt solid black;border-left:none;border-top:none;'>" + dt_d.Rows[v_index][1] + "</td>                                                          \n");
                        htmlStr.Append("				  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:none ;border-top:none;                                                                          \n");
                        htmlStr.Append("				  none'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("				  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:none;border-top:none;                                                                           \n");
                        htmlStr.Append("				  none'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("				  <td colspan=3 class=xl146 style='border-left:none;border-top:none;'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                                                                  \n");
                        htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
                        htmlStr.Append("				</tr>                                                                                                                                                                        \n");
                    
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
                    v_spacePerPage = 50;
                }
                //ESysLib.WriteLogError("v_spacePerPage Error:" + v_spacePerPage); 
                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("				 		<tr height=21 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                                                                      \n");
                            htmlStr.Append("						  <td height=21 class=xl65 style='height:" + v_rowHeightEmptyLast + "'></td>                                                                                                             \n");
                            htmlStr.Append("						  <td class=xl81 style=' border-bottom:1.0pt solid black;'>&nbsp;</td>                                                                                               \n");
                            htmlStr.Append("						  <td class=xl137 style=' border-bottom:1.0pt solid black;' ></td>                                                                                           \n");
                            htmlStr.Append("						  <td colspan=4 class=xl138 width=301 style='border-right:1.0pt solid black;border-top:none;border-bottom:1.0pt solid black;                                           \n");
                            htmlStr.Append("						  border-left:none;width:226pt'>&nbsp;</td>                                                                                                                \n");
                            htmlStr.Append("						  <td colspan=2 class=xl141 style='border-right:1.0pt solid black;border-left:none;border-top:none;border-bottom:1.0pt solid black;'></td>                   \n");
                            htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:1.0pt solid black;                                        \n");
                            htmlStr.Append("						  none'>&nbsp;</td>                                                                                                                                        \n");
                            htmlStr.Append("						  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:;border-top:none;border-bottom:1.0pt solid black;                                        \n");
                            htmlStr.Append("						  none'>&nbsp;</td>                                                                                                                                        \n");
                            htmlStr.Append("						  <td colspan=3 class=xl146 style='border-left:1.0pt solid black;;border-top:none;border-bottom:1.0pt solid black;border-right:1.0pt solid black;'>&nbsp;</td>\n");
                            htmlStr.Append("						  <td class=xl191 style=' border-right:1.0pt solid black;'>&nbsp;</td>                                                                                                \n");
                            htmlStr.Append("						</tr>                                                                                                                                                                \n");

                        }
                        else
                        {
                            htmlStr.Append("				<tr height=21 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                                                                              \n");
                            htmlStr.Append("				  <td height=21 class=xl65 style='height:" + v_rowHeightEmptyLast + "'></td>                                                                                                                      \n");
                            htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
                            htmlStr.Append("				  <td class=xl137 ></td>                                                                                                                                            \n");
                            htmlStr.Append("				  <td colspan=4 class=xl138 width=301 style='border-right:1.0pt solid black;border-top:none;                                                                                  \n");
                            htmlStr.Append("				  border-left:none;width:226pt'>&nbsp;</td>                                                                                                                        \n");
                            htmlStr.Append("				  <td colspan=2 class=xl141 style='border-right:1.0pt solid black;border-left:none;border-top:none;'></td>                                                          \n");
                            htmlStr.Append("				  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:none ;border-top:none;                                                                          \n");
                            htmlStr.Append("				  none'>&nbsp;</td>                                                                                                                                                \n");
                            htmlStr.Append("				  <td colspan=3 class=xl143 style='border-right:1.0pt solid black;border-left:none;border-top:none;                                                                           \n");
                            htmlStr.Append("				  none'>&nbsp;</td>                                                                                                                                                \n");
                            htmlStr.Append("				  <td colspan=3 class=xl146 style='border-left:none;border-top:none;'>&nbsp;</td>                                                                                  \n");
                            htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
                            htmlStr.Append("				</tr>                                                                                                                                                                        \n");

                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    htmlStr.Append("		 <tr height=18 style='height:" + (v_spacePerPage).ToString() + "pt'>                                                                                                                                                \n");
                    htmlStr.Append("		  <td height=18 class=xl65 style='height:" + (v_spacePerPage).ToString() + "pt'></td>                                                                                                                               \n");
                    htmlStr.Append("		  <td class=xl105>&nbsp;</td>                                                                                                                                                        \n");
                    htmlStr.Append("		  <td colspan=17 class=xl132 style='border-right:1.0pt solid black')>&nbsp;</td>                                                                                                      \n");
                    htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=2  style='height: 18pt'>                                                                                                                                                                \n");
                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");
                }


            }// for k                                                                                                                             

            htmlStr.Append("				 <tr height=18 style='height:15.2pt'>                                                                                                                                        \n");
            htmlStr.Append("				  <td height=18 class=xl65 style='height:15.2pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl88 style='border-top:1.0pt solid black '>&nbsp;</td>                                                                                                            \n");
            htmlStr.Append("				  <td class=xl89 colspan=2 style='mso-ignore:colspan'>C&#7897;ng ti&#7873;n                                                                                                  \n");
            htmlStr.Append("				  hàng <font class='font8'>Total</font><font class='font15'>:</font></td>                                                                                                    \n");
            htmlStr.Append("				  <td class=xl88 style='border-top:1.0pt solid black '>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl87 style='border-top:1.0pt solid black '>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl89>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl90>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=3 class=xl121 style='border-right:1.0pt solid black;border-left:                                                                                                \n");
            htmlStr.Append("				  none'>&nbsp;</td>                                                                                                                                                          \n");
            htmlStr.Append("				  <td class=xl89>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl89>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl92>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=3 class=xl121 style='border-left:none'>" + dt.Rows[0]["netamount_display"] + " &nbsp;</td>                                                                                                  \n");
            htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=18 style='height:15.2pt'>                                                                                                                                        \n");
            htmlStr.Append("				  <td height=18 class=xl65 style='height:15.2pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=3 rowspan=2 class=xl123 width=181 style='border-bottom:1.0pt solid black;                                                                                       \n");
            htmlStr.Append("				  width:136pt'>&nbsp;&nbsp;Tỷ giá: " + dt.Rows[0]["exchangerate"] + "</td>                                                                                                                                                   \n");
            htmlStr.Append("				  <td class=xl182 width=45 style='width:37.54pt'></td>                                                                                                                          \n");
            htmlStr.Append("				  <td rowspan=2 class=xl126 width=104 style='border-bottom:1.0pt solid black;                                                                                                 \n");
            htmlStr.Append("				  border-top:none;width:86.00pt'>&nbsp;</td>                                                                                                                                    \n");
            htmlStr.Append("				  <td class=xl93 colspan=2 style='mso-ignore:colspan'>Thu&#7871; VAT:<span                                                                                                   \n");
            htmlStr.Append("				  style='mso-spacerun:yes'> </span></td>                                                                                                                                     \n");
            htmlStr.Append("				  <td class=xl88>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl94>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl87>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=2 class=xl129 style='border-left:none'>" + dt.Rows[0]["TaxRate"] + " </td>                                                                                                       \n");
            htmlStr.Append("				  <td class=xl95>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=3 class=xl121 style='border-left:none'>" + dt.Rows[0]["vatamount_display"] + " &nbsp;</td>                                                                                                \n");
            htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=18 style='height:15.2pt'>                                                                                                                                        \n");
            htmlStr.Append("				  <td height=18 class=xl65 style='height:15.2pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl96 width=45 style='width:37.54pt'>&nbsp;</td>                                                                                                                     \n");
            htmlStr.Append("				  <td class=xl93 colspan=5 style='mso-ignore:colspan;border-right:1.0pt solid black'>T&#7893;ng                                                                               \n");
            htmlStr.Append("				  c&#7897;ng <font class='font8'>Grand total</font><font class='font15'>:</font></td>                                                                                        \n");
            htmlStr.Append("				  <td class=xl88 style='border-top:.5pt solid black'>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl88 style='border-top:.5pt solid black'>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl95>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=3 class=xl121 style='border-left:none'>" + dt.Rows[0]["totalamount_display"] + " &nbsp;</td>                                                                                         \n");
            htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=21 style='mso-height-source:userset;height:25.05pt'>                                                                                                             \n");
            htmlStr.Append("				  <td height=21 class=xl65 style='height:25.05pt'></td>                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl98 colspan=4 style='mso-ignore:colspan'>S&#7889; ti&#7873;n VND                                                                                                \n");
            htmlStr.Append("				  vi&#7871;t b&#7857;ng ch&#7919; <font class='font8'>(In words):</font><font                                                                                                \n");
            htmlStr.Append("				  class='font15'><span style='mso-spacerun:yes'> </span></font></td>                                                                                                         \n");
            htmlStr.Append("				  <td colspan=12 rowspan=2 class=xl130 width=477 style='border-bottom:1.0pt solid black;                                                                                      \n");
            htmlStr.Append("				  width:358pt'>" + read_prive + " &nbsp;</td>                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=13 style='mso-height-source:userset;height:10.05pt'>                                                                                                             \n");
            htmlStr.Append("				  <td height=13 class=xl65 style='height:10.05pt'></td>                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl81>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl88>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl88>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl88>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl88>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl191>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				                                                                                                                                                                             \n");
            htmlStr.Append("				 <tr height=18 style='mso-height-source:userset;height:20.8pt'>                                                                                                              \n");
            htmlStr.Append("				  <td height=18 class=xl65 style='height:13.8pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl72>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=6 class=xl73>NG&#431;&#7900;I MUA HÀNG/ <font class='font8'>BUYER</font></td>                                                                                  \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td colspan=6 class=xl73>NG&#431;&#7900;I BÁN HÀNG/ <font class='font8'>SELLER</font></td>                                                                                 \n");
            htmlStr.Append("				  <td class=xl78>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=18 style='height:15.84‬pt'>                                                                                                                                        \n");
            htmlStr.Append("				  <td height=18 class=xl65 style='height:15.84‬pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl72>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td colspan=6 class=xl73>Ký, ghi rõ h&#7885;, tên<span                                                                                                                     \n");
            htmlStr.Append("				  style='mso-spacerun:yes'> </span></td>                                                                                                                                     \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td colspan=6 class=xl73><span style='mso-spacerun:yes'> </span>Ký, &#273;óng                                                                                              \n");
            htmlStr.Append("				  d&#7845;u, ghi rõ h&#7885; tên<span style='mso-spacerun:yes'> </span></td>                                                                                                 \n");
            htmlStr.Append("				  <td class=xl78>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				 </tr> <!-- modify =================================================================================== -->                                                                   \n");
            htmlStr.Append("				                                                                                                                                                                             \n");
            htmlStr.Append("				 <tr class=xl97 height=25 style='mso-height-source:userset;height:10.6pt'>                                                                                                   \n");
            htmlStr.Append("				  <td height=25 class=xl97 style='height:10.6pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl164>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				  <td colspan=5 class=xl183><span style='mso-spacerun:yes'>   </span><font                                                                                                   \n");
            htmlStr.Append("				  class='font8'><span style='mso-spacerun:yes'> </span>(Sign, full name)</font></td>                                                                                         \n");
            htmlStr.Append("				  <td class=xl97></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl97></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl97></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl97></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl97></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td colspan=6 class=xl184>(Sign, stamp, full name)</td>                                                                                                                    \n");
            htmlStr.Append("				  <td class=xl168>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				                                                                                                                                                                             \n");
            htmlStr.Append("				                                                                                                                                                                             \n");
            htmlStr.Append("				  <tr class=xl97 height=25 style='mso-height-source:userset;height:15.6pt'>                                                                                                  \n");
            htmlStr.Append("					  <td height=25 class=xl97 style='height:15.6pt'></td>                                                                                                                   \n");
            htmlStr.Append("					  <td class=xl164>&nbsp;</td>                                                                                                                                            \n");
            htmlStr.Append("					  <td colspan=5 class=xl183><span style='mso-spacerun:yes'>   </td>                                                                                                      \n");
            htmlStr.Append("					  <td class=xl97></td>                                                                                                                                                   \n");
            htmlStr.Append("					  <td class=xl97></td>                                                                                                                                                   \n");
            htmlStr.Append("					  <td class=xl97></td>                                                                                                                                                   \n");
            htmlStr.Append("					  <td class=xl97></td>                                                                                                                                                   \n");
            htmlStr.Append("					  <td class=xl97></td>                                                                                                                                                   \n");
            htmlStr.Append("					  <td colspan=6 class=xl184></td>                                                                                                                                        \n");
            htmlStr.Append("					  <td class=xl168>&nbsp;</td>                                                                                                                                            \n");
            htmlStr.Append("				  </tr>                                                                                                                                                                      \n");
            htmlStr.Append("				                                                                                                                                                                             \n");
            htmlStr.Append("				 <!-- modify =================================================================================== -->                                                                         \n");
            htmlStr.Append("				 <tr height=21 style='mso-height-source:userset;height:15.75pt'>                                                                                                             \n");
            htmlStr.Append("				  <td height=21 class=xl65 style='height:15.75pt'></td>                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl72>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
           if(dt.Rows[0]["sign_yn"] == "Y")
            {
                htmlStr.Append("				  <td class=xl99 colspan=3 style='mso-ignore:colspan'><span style='mso-ignore:vglayout;position:                                                                             \n");
                htmlStr.Append("				  absolute;z-index:4;margin-left:130px;margin-top:9px;width:62px;height:42px'><img                                                                                           \n");
                htmlStr.Append("				  width=62 height=42  src='D:/webproject/e-invoice-ws/02.Web/EInvoice/img/check_signed.png'  v:shapes='Rectangle_x0020_4'></span>Signature Valid</td>                         \n");

            }
            else
            {
                htmlStr.Append("				  <td class=xl99 colspan=3 style='mso-ignore:colspan'></td>                         \n");

            }

            htmlStr.Append("				  <td class=xl83>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl83>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl83>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				  <td class=xl101>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				  <td class=xl75>&nbsp;</td>                                                                                                                                                 \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				                                                                                                                                                                             \n");
            htmlStr.Append("		 <tr height=19 style='mso-height-source:userset;height:14.4pt'>                                                                                                                      \n");
            htmlStr.Append("		  <td height=19 class=xl65 style='height:14.4pt'></td>                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl72>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		                                                                                                                                                                                     \n");
            htmlStr.Append("		  <td class=xl97 colspan=2 style='mso-ignore:colspan'>Mã CQT:                                                                                                            \n");
            htmlStr.Append("		  <span style='mso-spacerun:yes'> </span></td>                                                                                                                         \n");
            htmlStr.Append("		                                                                                                                                                                                     \n");
            htmlStr.Append("		  <td colspan=6 rowspan=2 class=xl108 width=302 style='width:226pt'>" + dt.Rows[0]["cqt_mccqt_id"] + " </td>                                                                                                \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl102 colspan=2 style='mso-ignore:colspan'>&#272;&#432;&#7907;c ký                                                                                                       \n");
            htmlStr.Append("		  b&#7903;i<font class='font18'>:<span style='mso-spacerun:yes'> </span></font></td>                                                                                                 \n");
            htmlStr.Append("		  <td colspan=5  class=xl135 width=198 style='border-right:1.0pt solid black;                                                                                                         \n");
            htmlStr.Append("		  width:149pt'>" + dt.Rows[0]["SignedBy"] + "</td>                                                                                                                                                    \n");
            htmlStr.Append("		  <td class=xl75>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		 </tr>                                                                                                                                                                               \n");
            htmlStr.Append("		                                                                                                                                                                                     \n");
            htmlStr.Append("		                                                                                                                                                                                     \n");
            htmlStr.Append("		 <tr height=22 style='mso-height-source:userset;height:16.5pt'>                                                                                                                      \n");
            htmlStr.Append("		  <td height=22 class=xl65 style='height:16.5pt'></td>                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl72>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		  <td class=xl97 colspan=6 style='mso-ignore:colspan'>Tra c&#7913;u t&#7841;i                                                                                                        \n");
            htmlStr.Append("		  website: " + dt.Rows[0]["WEBSITE_EI"] + "</td>                                                                                                                                 \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl103 colspan=2 style='mso-ignore:colspan'>Ngày Ký:</td>                                                                                                                 \n");
            htmlStr.Append("		  <td class=xl104 colspan=3 style='mso-ignore:colspan'>" + dt.Rows[0]["SignedDate"] + "</td>                                                                                                               \n");
            htmlStr.Append("		  <td class=xl88>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		  <td class=xl87>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		  <td class=xl75>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		 </tr>                                                                                                                                                                               \n");

            htmlStr.Append("		 <tr height=17 style='mso-height-source:userset;height:13.05pt'>                                                                                                                     \n");
            htmlStr.Append("		  <td height=17 class=xl65 style='height:13.05pt'></td>                                                                                                                              \n");
            htmlStr.Append("		  <td class=xl72>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65>Mã nhận hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "</td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl65></td>                                                                                                                                                               \n");
            htmlStr.Append("		  <td class=xl75>&nbsp;</td>                                                                                                                                                         \n");
            htmlStr.Append("		 </tr>                                                                                                                                                                 \n");
            htmlStr.Append("				 <tr height=18 style='height:15.84‬pt'>                                                                                                                                        \n");
            htmlStr.Append("				  <td height=18 class=xl65 style='height:15.84‬pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl105>&nbsp;</td>                                                                                                                                                \n");
            htmlStr.Append("				  <td colspan=17 class=xl132 style='border-right:1.0pt solid black'>( C&#7847;n                                                                                               \n");
            htmlStr.Append("				  ki&#7875;m tra, &#273;&#7889;i chi&#7871;u khi l&#7853;p, giao nh&#7853;n hoá                                                                                              \n");
            htmlStr.Append("				  &#273;&#417;n )</td>                                                                                                                                                       \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=10 style='mso-height-source:userset;height:9.36pt'>                                                                                                               \n");
            htmlStr.Append("				  <td height=10 class=xl65 style='height:9.36pt'></td>                                                                                                                        \n");
            htmlStr.Append("				  <td class=xl65></td>                                                                                                                                                       \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				  <td class=xl106></td>                                                                                                                                                      \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=22 style='mso-height-source:userset;height:20.7‬pt'>                                                                                                             \n");
            htmlStr.Append("				  <td height=22 class=xl65 style='height:20.7‬pt'></td>                                                                                                                      \n");
            htmlStr.Append("				  <td colspan=18 class=xl133>BK VINA CO., LTD</td>                                                                                                                           \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <tr height=18 style='height:15.84‬pt'>                                                                                                                                        \n");
            htmlStr.Append("				  <td height=18 class=xl65 style='height:15.84‬pt'></td>                                                                                                                       \n");
            htmlStr.Append("				  <td colspan=18 class=xl134>Lot A-5C-CN, My Phuoc 3 I.P, Chanh Phu Hoa Ward,                                                                                                \n");
            htmlStr.Append("				  Ben Cat Town, Binh Duong Province<span style='mso-spacerun:yes'>  </span>!!!!!                                                                                                 \n");
            htmlStr.Append("				  Tel:0274-3559-826 ~ 30 !!!!! Fax: 0274 - 3559 831</td>                                                                                                                         \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <![if supportMisalignedColumns]>                                                                                                                                            \n");
            htmlStr.Append("				 <tr height=0 style='display:none'>                                                                                                                                          \n");
            htmlStr.Append("				  <td width=3 style='width:2.65pt'></td>                                                                                                                                        \n");
            htmlStr.Append("				  <td width=5 style='width:5.3pt'></td>                                                                                                                                        \n");
            htmlStr.Append("				  <td width=29 style='width:28.00pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=86 style='width:86.00pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=66 style='width:64.84pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=45 style='width:37.54pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=104 style='width:86.00pt'></td>                                                                                                                                     \n");
            htmlStr.Append("				  <td width=22 style='width:22.50pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=51 style='width:50.27pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=14 style='width:13.23pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=17 style='width:17.20pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=36 style='width:35.72pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=35 style='width:32.76pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=20 style='width:19.85pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=35 style='width:32.76pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=22 style='width:22.50pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=53 style='width:52.92pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=68 style='width:61.2pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				  <td width=22 style='width:22.50pt'></td>                                                                                                                                      \n");
            htmlStr.Append("				 </tr>                                                                                                                                                                       \n");
            htmlStr.Append("				 <![endif]>                                                                                                                                                                  \n");
            htmlStr.Append("				</table>                                                                                                                                                                     \n");
            htmlStr.Append("   </ body >          \n");
            htmlStr.Append("    </ html >               \n");

            string filePath = "";

            /*using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\GIT\EInvoice\PDF\" + tei_einvoice_m_pk + ".html"))
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

        public static void Genneral_Barcode(string txtBarcode)
        {
            string fileName = "D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\Barcode_Code.png";
            string fileName_e = "D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\Barcode_Code.gif";
            // string fileName     = "C:/Users/genuwin/Desktop/Barcode_Code.png";
            // string fileName_e   = "C:/Users/genuwin/Desktop/Barcode_Code.gif";
            /*Bitmap bm = new Bitmap(fileName);
            // if already on disk, use it. otherwise create it.
            // context.Response.ContentType = "image/gif";
            if (txtBarcode.Length > 0)
            {
                Barcode128 code128 = new Barcode128();
                code128.CodeType = Barcode.CODE128;
                code128.ChecksumText = true;
                code128.GenerateChecksum = true;
                code128.StartStopText = true;
                code128.Code = txtBarcode;
                //bm = new Bitmap(code128.CreateDrawingImage(Color.Black, Color.White));
                bm.Save(fileName_e); // to disk
            }*/


            Bitmap oBmp1 = new Bitmap(470, 50);
            Graphics oGrp1 = Graphics.FromImage(oBmp1);
            //Graphics oGrp2 = Graphics.FromImage(oBmp1);

            //string sText = "*ENG071121010101S*";
            //string sText2 = "ENG071121010101S";

            SolidBrush oBrush = new SolidBrush(Color.White);
            SolidBrush oBrushWrite = new SolidBrush(Color.Black);

            oGrp1.FillRectangle(oBrush, 0, 0, 470, 100);

            PrivateFontCollection pfc = new PrivateFontCollection();
            pfc.AddFontFile("D:\\webproject\\e-invoice-ws\\font\\free3of9.ttf");
            FontFamily family = new FontFamily("Free 3 of 9", pfc);

            Font oFont = new Font(family, 45);
            //Font oFont2 = new Font("Arial", 16);

            PointF oPoint = new PointF(5F, 5F);
            //PointF oPoint2 = new PointF(135, 60);

            oGrp1.DrawString(txtBarcode, oFont, oBrushWrite, oPoint);
            //oGrp2.DrawString(sText2, oFont2, oBrushWrite, oPoint2);

            //Response.ContentType = "image/jpeg";
            oBmp1.Save(fileName, ImageFormat.Jpeg);

        }
    }
}