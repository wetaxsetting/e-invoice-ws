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
    public class KPX
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


            htmlStr.Append("    <!DOCTYPE html PUBLIC ' -//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>                                                               \n");
            htmlStr.Append("    <html>                                                               \n");
            htmlStr.Append("    <head>                                                               \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                                                                  \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("    <script type='text/javascript'                                                               \n");
            htmlStr.Append("    	src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                               \n");
            htmlStr.Append("    <title>Report E-Invoice</title>                                                               \n");
            htmlStr.Append("    <!-- Normalize or reset CSS with your favorite library -->                                                               \n");
            // htmlStr.Append("    <link rel='stylesheet'                                                               \n");
            //htmlStr.Append("    	href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'>                                                               \n");
            // htmlStr.Append("                                                                   \n");
            htmlStr.Append("    <!-- Load paper.css for happy printing -->                                                               \n");
            //htmlStr.Append("    <link rel='stylesheet'                                                               \n");
            //htmlStr.Append("    	href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>                                                               \n");
            //htmlStr.Append("                                                                   \n");
            htmlStr.Append("    <!-- Set page size here: A5, A4 or A3 -->                                                               \n");
            htmlStr.Append("    <!-- Set also 'landscape' if you need -->                                                               \n");
            htmlStr.Append("    <style>                                                               \n");
            htmlStr.Append("    @page {                                                               \n");
            htmlStr.Append("    	size: A4;                                                               \n");
            htmlStr.Append("    	padding-left: 50px;                                                               \n");
            htmlStr.Append("    }                                                               \n");
            htmlStr.Append("    </style>                                                               \n");
            //htmlStr.Append("    <link href='https://fonts.googleapis.com/css?family=Tangerine:700'                                                               \n");
            //htmlStr.Append("    	rel='stylesheet' type='text/css'>                                                               \n");
            htmlStr.Append("    <style>                                                               \n");
            htmlStr.Append("    /*body   { font-family: serif }                                                               \n");
            htmlStr.Append("        h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                               \n");
            htmlStr.Append("        h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                               \n");
            htmlStr.Append("        h4     { font-size: 13pt; line-height: 1mm }                                                               \n");
            htmlStr.Append("        h2 + p { font-size: 18pt; line-height: 7mm }                                                               \n");
            htmlStr.Append("        h3 + p { font-size: 14pt; line-height: 7mm }                                                               \n");
            htmlStr.Append("        li     { font-size: 11pt; line-height: 5mm }                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("        h1      { margin: 0 }                                                               \n");
            htmlStr.Append("        h1 + ul { margin: 2mm 0 5mm }                                                               \n");
            htmlStr.Append("        h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                               \n");
            htmlStr.Append("        h2 + p,                                                               \n");
            htmlStr.Append("        h3 + p  { margin: 0 0 3mm 50mm }                                                               \n");
            htmlStr.Append("        //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                               \n");
            htmlStr.Append("        h4 + ul { margin: 5mm 0 0 50mm }                                                               \n");
            htmlStr.Append("        article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                               \n");
            htmlStr.Append("    body {                                                               \n");
            htmlStr.Append("    	color: blue;                                                               \n");
            htmlStr.Append("    	font-size: 100%;                                                               \n");
            htmlStr.Append("    	background-image: url('assets/Solution.jpg');                                                               \n");
            htmlStr.Append("    }                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("    h1 {                                                               \n");
            htmlStr.Append("    	color: #00FF00;                                                               \n");
            htmlStr.Append("    }                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("    p {                                                               \n");
            htmlStr.Append("    	color: rgb(0, 0, 255)                                                               \n");
            htmlStr.Append("    }                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("    headline1 {                                                               \n");
            htmlStr.Append("    	background-image: url(assets/Solution.jpg);                                                               \n");
            htmlStr.Append("    	background-repeat: no-repeat;                                                               \n");
            htmlStr.Append("    	background-position: left top;                                                               \n");
            htmlStr.Append("    	padding-top: 68px;                                                               \n");
            htmlStr.Append("    	margin-bottom: 50px;                                                               \n");
            htmlStr.Append("    }                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("    headline2 {                                                               \n");
            htmlStr.Append("    	background-image: url(images/newsletter_headline2.gif);                                                               \n");
            htmlStr.Append("    	background-repeat: no-repeat;                                                               \n");
            htmlStr.Append("    	background-position: left top;                                                               \n");
            htmlStr.Append("    	padding-top: 68px;                                                               \n");
            htmlStr.Append("    	padding-left: 68px;                                                               \n");
            htmlStr.Append("    }                                                               \n");

            htmlStr.Append("     <!--table																																\n");
            htmlStr.Append("     	{mso-displayed-decimal-separator:'\\.';                                                                                              \n");
            htmlStr.Append("     	mso-displayed-thousand-separator:'\\,';}                                                                                             \n");
            htmlStr.Append("     .font56926                                                                                                                             \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                  \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font66926                                                                                                                             \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                  \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font76926                                                                                                                             \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                  \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font86926                                                                                                                             \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                  \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font96926                                                                                                                             \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                  \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font106926                                                                                                                            \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                  \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font116926                                                                                                                            \n");
            htmlStr.Append("     	{color:#0066CC;                                                                                                                     \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font126926                                                                                                                            \n");
            htmlStr.Append("     	{color:red;                                                                                                                         \n");
            htmlStr.Append("     	font-size:8.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font136926                                                                                                                            \n");
            htmlStr.Append("     	{color:red;                                                                                                                         \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font146926                                                                                                                            \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                  \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font156926                                                                                                                            \n");
            htmlStr.Append("     	{color:black;                                                                                                                       \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .font166926                                                                                                                            \n");
            htmlStr.Append("     	{color:#0066CC;                                                                                                                     \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                \n");
            htmlStr.Append("     .xl656926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl656926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl6669261                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl676926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl686926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl696926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl706926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl716926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl726926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl736926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl746926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl7469261                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl756926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl766926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl776926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl786926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl796926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl806926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl816926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl826926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl836926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl846926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl856926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl866926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl876926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl886926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl896926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl906926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl916926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl926926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl936926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl946926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl956926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#C00000;                                                                                                                      \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl966926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl976926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl986926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl996926                                                                                                                              \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1006926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1016926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1026926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1036926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1046926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1056926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1066926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#0070C0;                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1076926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1086926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1096926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1106926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1116926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1126926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1136926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1146926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1156926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1166926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1176926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1186926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1196926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1206926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1216926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1226926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1236926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1246926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1256926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1266926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:0%;                                                                                                               \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1276926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1286926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:black;                                                                                                                        \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1296926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:black;                                                                                                                        \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1306926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1316926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1326926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1336926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:11.25pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1346926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1356926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1366926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1376926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1386926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\#\\,\\#\\#0';                                                                                                      \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1396926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\#\\,\\#\\#0';                                                                                                      \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1406926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1416926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1426926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1436926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1446926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1456926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1466926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1476926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1486926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1496926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'Short Date';                                                                                                     \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1506926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'Short Date';                                                                                                     \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1516926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1526926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1536926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:20.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1546926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1556926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1566926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:20.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1576926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:general;                                                                                                                 \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1586926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1596926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1606926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1616926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1626926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:#002060;                                                                                                                      \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                      \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                         \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                   \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1636926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1646926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:17.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1656926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:14.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1666926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1676926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1686926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1696926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1706926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1716926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1726926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1736926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1746926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1756926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1766926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1776926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1786926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1796926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1806926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1816926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1826926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1836926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1846926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:red;                                                                                                                          \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.75pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1856926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1866926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1876926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1886926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1896926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1906926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1916926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1926926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1936926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1946926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:11.25pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl1956926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1966926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                             \n");
            htmlStr.Append("     	text-align:right;                                                                                                                   \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.75pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl1976926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:11.25pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1986926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:11.5pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl1996926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                \n");
            htmlStr.Append("     .xl2006926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     .xl2016926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:black;                                                                                                                        \n");
            htmlStr.Append("     	font-size:8.0pt;                                                                                                                    \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl2026926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:black;                                                                                                                        \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("     .xl2036926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:black;                                                                                                                        \n");
            htmlStr.Append("     	font-size:13.75pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:left;                                                                                                                    \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:none;                                                                                                                    \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                 \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                      \n");
            htmlStr.Append("      .xl6569261 {												                                                                            \n");
            htmlStr.Append("          padding: 0px;                                                                                                                     \n");
            htmlStr.Append("          mso-ignore: padding;                                                                                                              \n");
            htmlStr.Append("          color: windowtext;                                                                                                                \n");
            htmlStr.Append("          font-size: 12.50pt;                                                                                                                \n");
            htmlStr.Append("          font-weight: 400;                                                                                                                 \n");
            htmlStr.Append("          font-style: normal;                                                                                                               \n");
            htmlStr.Append("          text-decoration: none;                                                                                                            \n");
            htmlStr.Append("          font-family: 'Times New Roman', serif;                                                                                            \n");
            htmlStr.Append("          mso-font-charset: 0;                                                                                                              \n");
            htmlStr.Append("          mso-number-format: General;                                                                                                       \n");
            htmlStr.Append("          text-align: general;                                                                                                              \n");
            htmlStr.Append("          vertical-align: middle;                                                                                                           \n");
            htmlStr.Append("          background: white;                                                                                                                \n");
            htmlStr.Append("          mso-pattern: black none;                                                                                                          \n");
            htmlStr.Append("          white-space: nowrap;                                                                                                              \n");
            htmlStr.Append("      }                                                                                                                                     \n");
            htmlStr.Append("     .xl2046926                                                                                                                             \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                       \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                 \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                   \n");
            htmlStr.Append("     	font-size:15.0pt;                                                                                                                   \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                    \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                  \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                               \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                 \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                          \n");
            htmlStr.Append("     	text-align:center;                                                                                                                  \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                              \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("     	border-right:none;                                                                                                                  \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                 \n");
            htmlStr.Append("     	border-left:none;                                                                                                                   \n");
            htmlStr.Append("     	background:white;                                                                                                                   \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                             \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                \n");
            htmlStr.Append("     -->                                                                                                                                    \n"); htmlStr.Append("    </style>                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("    </head>                                                               \n");
            htmlStr.Append("    <body class='A4'>                                                               \n");
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

            double v_totalHeightLastPage = 223.5;// 243.5

            double v_totalHeightPage = 530;//   518;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 510;// 530;

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

                htmlStr.Append("     <table border=0 cellpadding=0 cellspacing=0 width=716 class=xl656926													    \n");
                htmlStr.Append("      style='border-collapse:collapse;table-layout:fixed;width:539pt'>                                                          \n");
                htmlStr.Append("      <col class=xl656926 width=6 style='mso-width-source:userset;mso-width-alt:                                                \n");
                htmlStr.Append("      219;width:6.25pt'>                                                                                                           \n");
                htmlStr.Append("      <col class=xl656926 width=32 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      1170;width:30pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=67 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      2450;width:87.5pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=53 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      1938;width:25pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=39 span=2 style='mso-width-source:userset;                                                      \n");
                htmlStr.Append("      mso-width-alt:1426;width:36.25pt'>                                                                                           \n");
                htmlStr.Append("      <col class=xl656926 width=67 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      2450;width:62.5pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=11 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      402;width:10pt'>                                                                                                           \n");
                htmlStr.Append("      <col class=xl656926 width=75 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      2742;width:70pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=54 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      1974;width:51.25pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=29 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      1060;width:22pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=6 style='mso-width-source:userset;mso-width-alt:                                                \n");
                htmlStr.Append("      219;width:6.25pt'>                                                                                                           \n");
                htmlStr.Append("      <col class=xl656926 width=39 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      1426;width:36.25pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=46 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      1682;width:35pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=6 style='mso-width-source:userset;mso-width-alt:                                                \n");
                htmlStr.Append("      219;width:6.25pt'>                                                                                                           \n");
                htmlStr.Append("      <col class=xl656926 width=39 style='mso-width-source:userset;mso-width-alt:                                               \n");
                htmlStr.Append("      1426;width:46.25pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=102 style='mso-width-source:userset;mso-width-alt:                                              \n");
                htmlStr.Append("      3730;width:86.25pt'>                                                                                                         \n");
                htmlStr.Append("      <col class=xl656926 width=6 style='mso-width-source:userset;mso-width-alt:                                                \n");
                htmlStr.Append("      219;width:6.25pt'>                                                                                                           \n");
                htmlStr.Append("     <tr height=33 style='mso-height-source:userset;height:31.2pt'>                                                            \n");
                htmlStr.Append("      <td height=33 width=6 style='height:31.2pt;width:4pt' align=left class=xl1116926 valign=top><![if !vml]><span style='mso-ignore:vglayout;					  \n");
                htmlStr.Append("      position:absolute;z-index:2;margin-left:1px;margin-top:31px;width:185px;                                                                                        \n");
                htmlStr.Append("      height:83px'><img width=185 height=83                                                                                                                           \n");
                htmlStr.Append("      src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\KPX_001.png'                                                                                              \n");
                htmlStr.Append("      v:shapes='Picture_x0020_6'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                                 \n");
                htmlStr.Append("      <table cellpadding=0 cellspacing=0>                                                                                                                             \n");
                htmlStr.Append("       <tr>                                                                                                                                                           \n");
                htmlStr.Append("        <td height=33  width=6 style='height:31.2pt;width:6.25pt'>&nbsp;</td>                                                                                           \n");
                htmlStr.Append("       </tr>                                                                                                                                                          \n");
                htmlStr.Append("      </table>                                                                                                                                                        \n");
                htmlStr.Append("      </span></td>                                                                                                                                                    \n");
                htmlStr.Append("      <td class=xl676926 width=32 style='width:30pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl986926 width=67 style='width:62.5pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl986926 width=53 style='width:40pt'>&nbsp;</td>                                                                                                      \n");
                if(dt.Rows[0]["tei_company_pk"].ToString()== "662")
                {
                    htmlStr.Append("      <td class=xl1566926 colspan=2 width=78 style='width:58pt'>" + dt.Rows[0]["Seller_Name"] + "</td>                                                                               \n");

                }
                else
                {
                    htmlStr.Append("       <td class=xl1566926 colspan=2 width=78 style='width:58pt;font-size:13.5pt'>" + dt.Rows[0]["Seller_Name"] + "</td>                                                             \n");

                }
                  htmlStr.Append("      <td class=xl1426926 width=67 style='width:62.5pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl1426926 width=11 style='width:10pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl1426926 width=75 style='width:70pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl1426926 width=54 style='width:51.25pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl1426926 width=29 style='width:22pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl1426926 width=6 style='width:6.25pt'>&nbsp;</td>                                                                                                       \n");
                htmlStr.Append("      <td class=xl1426926 width=39 style='width:36.25pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl986926 width=46 style='width:35pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl986926 width=6 style='width:6.25pt'>&nbsp;</td>                                                                                                        \n");
                htmlStr.Append("      <td class=xl986926 width=39 style='width:36.25pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl986926 width=102 style='width:96.25pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl996926 width=6 style='width:6.25pt'>&nbsp;</td>                                                                                                        \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=24 class=xl1416926 style='mso-height-source:userset;height:22.5pt'>                                                                                   \n");
                htmlStr.Append("      <td height=24 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926 colspan=2>Mã s&#7889; thu&#7871; (<font class='font106926'>Tax code</font><font class='font76926'>)</font></td>                              \n");
                htmlStr.Append("      <td class=xl1436926></td>                                                                                                                                       \n");
                htmlStr.Append("      <td class=xl1456926>:&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("      <td colspan=9 class=xl1586926 width=396 style='width:299pt'>  " + dt.Rows[0]["Seller_TaxCode"] + "</td>                                                                      \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=33 style='mso-height-source:userset;height:31.2pt'>                                                                                                  \n");
                htmlStr.Append("      <td height=33 class=xl1416926 style='height:31.2pt'>&nbsp;</td>                                                                                                \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926 colspan=2>&#272;&#7883;a ch&#7881; (<font                                                                                                    \n");
                htmlStr.Append("      class='font106926'>Ad</font><font                                                                                                                               \n");
                htmlStr.Append("      class='font106926'>dress</font><font class='font76926'>)</font></td>                                                                                            \n");
                htmlStr.Append("      <td class=xl1446926 width=67 style='width:62.5pt'></td>                                                                                                           \n");
                htmlStr.Append("      <td class=xl1466926 width=11 style='width:10pt'>:&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td colspan=9 class=xl1596926 width=396 style='width:299pt'>" + dt.Rows[0]["SELLER_ADDRESS"] + "</td>                                                                        \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=20 style='mso-height-source:userset;height:19.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=20 class=xl1416926 style='height:15.6pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926 colspan=2>&#272;i&#7879;n tho&#7841;i (<font class='font96926'>Tel)</font></td>                                                              \n");
                htmlStr.Append("      <td class=xl1436926 dir=LTR></td>                                                                                                                               \n");
                htmlStr.Append("      <td class=xl1476926>:&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("      <td colspan=5 class=xl1606926 width=203 style='width:153pt'>" + dt.Rows[0]["Seller_Tel"] + "</td>                                                                            \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl716926 width=6 style='width:6.25pt'>&nbsp;</td>                                                                                                        \n");
                htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl716926 width=102 style='width:96.25pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl726926 width=6 style='width:6.25pt'>&nbsp;</td>                                                                                                        \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=20 style='mso-height-source:userset;height:15.6pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=20 class=xl1416926 style='height:15.6pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926 colspan=2>S&#7889; tài kho&#7843;n (<font class='font106926'>Acc. code</font><font class='font76926'>)</font></td>                           \n");
                htmlStr.Append("      <td class=xl1436926 dir=LTR></td>                                                                                                                               \n");
                htmlStr.Append("      <td class=xl1476926>:&nbsp;</td>                                                                                                                                \n");
                htmlStr.Append("      <td colspan=9 class=xl1606926 width=396 style='width:299pt;font-size: 12.0pt'>                                                                                  \n");
                htmlStr.Append("      " + dt.Rows[0]["SELLER_ACCOUNTNO"] + "  " + dt.Rows[0]["BANK_NM78"] + "</td>                                                                                                               \n");
                htmlStr.Append("      <td class=xl726926 width=6 style='width:6.25pt'>&nbsp;</td>                                                                                                        \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("                                                                                                                                                                      \n");
                htmlStr.Append("     <tr height=6 style='mso-height-source:userset;height:2.1pt'>                                                                                                     \n");
                htmlStr.Append("      <td colspan=2 height=6 class=xl666926 style='height:2.1pt;border-top:.5pt solid windowtext;border-left:.5pt solid windowtext;'>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=9 class=xl1626926 dir=LTR>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl686926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=13 style='mso-height-source:userset;height:9.95pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=13 class=xl1416926 style='height:9.95pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td colspan=16 class=xl1636926>Mã c&#7911;a C&#417; quan thu&#7871;: <font                                                                                      \n");
                htmlStr.Append("      class='font166926'>" + dt.Rows[0]["cqt_mccqt_id"] + "</font></td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl726926 width=6 style='width:6.25pt'>&nbsp;</td>                                                                                                        \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=6 style='mso-height-source:userset;height:1pt'>                                                                                                       \n");
                htmlStr.Append("      <td colspan=2 height=6 class=xl1416926 style='height:1pt;border-bottom:.5pt solid windowtext;border-left:.5pt solid windowtext;'>&nbsp;</td>                                                                                           \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=9 class=xl1616926 dir=LTR>&nbsp;</td>                                                                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=4 style='mso-height-source:userset;height:1pt'>                                                                                                       \n");
                htmlStr.Append("      <td height=4 class=xl666926 style='height:1pt;border-left:.5pt solid windowtext;'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl736926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl736926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl736926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl676926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl676926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl676926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl676926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl676926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl676926 style='border-top:none'>&nbsp;</td>                                                                                                          \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl686926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=24 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=10 class=xl1646926>HÓA &#272;&#416;N GIÁ TR&#7882; GIA T&#258;NG</td>                                                                               \n");
                htmlStr.Append("      <td class=xl696926 colspan=3><font class='font106926'></font><font                                                                                              \n");
                htmlStr.Append("      class='font86926'></font></td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl1576926></td>                                                                                                                                       \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=24 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=10 class=xl1656926>(VAT INVOICE)</td>                                                                                                               \n");
                htmlStr.Append("      <td class=xl696926 colspan=3>Ký hi&#7879;u (<font class='font106926'>Serial</font><font class='font106926'>)</span> </td>                                                                                              \n");
                htmlStr.Append("      <td class=xl1576926>: " + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</td>                                                                                                                   \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=21 style='mso-height-source:userset;height:15.95pt'>                                                                                                  \n");
                htmlStr.Append("      <td height=21 class=xl1416926 style='height:15.95pt'>&nbsp;</td>                                                                                                \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=10 class=xl1996926></td>                                                                                                                            \n");
                htmlStr.Append("      <td class=xl696926 colspan=2>S&#7889; <font class='font106926'>(No</font><font                                                                                  \n");
                htmlStr.Append("      class='font86926'>)</font></td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl1486926><font class='font156926'>:</font><font class='font136926'><span                                                                             \n");
                htmlStr.Append("      style='mso-spacerun:yes'> " + dt.Rows[0]["InvoiceNumber"] + "</span></font></td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=24 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=10 class=xl1666926 width=412 style='width:309pt'>Ngày <font                                                                                         \n");
                htmlStr.Append("      class='font106926'>(Date)</font><font class='font96926'><span                                                                                                   \n");
                htmlStr.Append("      style='mso-spacerun:yes'> " + dt.Rows[0]["invoiceissueddate_dd"] + " </span></font><font class='font76926'><span                                                                             \n");
                htmlStr.Append("      style='mso-spacerun:yes'></span>tháng </font><font class='font106926'>(month)</font><font                                                                       \n");
                htmlStr.Append("      class='font76926'><span style='mso-spacerun:yes'> " + dt.Rows[0]["invoiceissueddate_mm"] + " </span>n&#259;m </font><font                                                                    \n");
                htmlStr.Append("      class='font106926'>(year)</font><font class='font76926'><span                                                                                                   \n");
                htmlStr.Append("      style='mso-spacerun:yes'> " + dt.Rows[0]["invoiceissueddate_yyyy"] + " </span></font></td>                                                                                                   \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                             \n");
                htmlStr.Append("     <tr height=4 style='mso-height-source:userset;height:1pt'>                                                                                                       \n");
                htmlStr.Append("      <td height=4 class=xl1416926 style='height:1pt;border-bottom:.5pt solid windowtext;border-left:.5pt solid windowtext;'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926 width=39 style='width:36.25pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl656926 width=39 style='width:36.25pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl656926 width=67 style='width:62.5pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=4 style='mso-height-source:userset;height:1pt'>                                                                                                       \n");
                htmlStr.Append("      <td height=4 class=xl666926 style='height:1pt;border-left:.5pt solid windowtext'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl756926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl756926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl756926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl676926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=11 class=xl1676926 style='border-right:.5pt solid black'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=24 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926 colspan=6>H&#7885; tên ng&#432;&#7901;i mua hàng <font                                                                                       \n");
                htmlStr.Append("      class='font106926'>(Customer's name)</font><font class='font76926'>:</font>" + dt.Rows[0]["buyer"] + "</td>                                                              \n");
                htmlStr.Append("      <td colspan=11 class=xl1696926 style='border-right:.5pt solid black'>&nbsp;</td>                                                                                \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=29 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=29 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926 colspan=4>Tên &#273;&#417;n v&#7883; <font                                                                                                   \n");
                htmlStr.Append("      class='font106926'>(Company's name)</font><font class='font76926'>:</font></td>                                                                \n");
                //htmlStr.Append("      <td class=xl766926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td colspan=12 class=xl1706926>" + dt.Rows[0]["buyerlegalname"] + "&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("      <td class=xl1406926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=24 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926 colspan=3>Mã s&#7889; thu&#7871; <font class='font106926'>(Tax                                                                               \n");
                htmlStr.Append("      code)</font><font class='font76926'>:<span style='mso-spacerun:yes'> </span></font>" + dt.Rows[0]["BuyerTaxCode"] + "</td>                                                         \n");
                htmlStr.Append("      <td class=xl766926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl766926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1406926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=24 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=24 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926 colspan=2>&#272;&#7883;a ch&#7881; <font class='font106926'>(Address)</font><font                                                            \n");
                htmlStr.Append("      class='font76926'>:<span style='mso-spacerun:yes'> </span></font></td>                                                                                          \n");
                htmlStr.Append("      <td class=xl7469261 colspan=14>" + dt.Rows[0]["BuyerAddress"] + "</td>                                                                                                       \n");
                htmlStr.Append("      <!-- <td class=xl766926>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("      <td class=xl1126926>&nbsp;</td> -->                                                                                                                             \n");
                htmlStr.Append("      <td class=xl1406926>&nbsp;</td>                                                                                                                                 \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr height=26 style='mso-height-source:userset;height:22.5pt'>                                                                                                   \n");
                htmlStr.Append("      <td height=26 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("      <td class=xl746926 colspan=6>Hình th&#7913;c thanh toán <font                                                                                                   \n");
                htmlStr.Append("      class='font106926'>(Payment Method)</font><font class='font76926'>:<span                                                                                        \n");
                htmlStr.Append("      style='mso-spacerun:yes'> " + dt.Rows[0]["PaymentMethodCK"] + " </span></font></td>                                                                                             \n");
                htmlStr.Append("      <td class=xl776926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl746926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");

                htmlStr.Append("   <tr height=26 style='mso-height-source:userset;height:22.5pt'>											   \n");
                htmlStr.Append("     <td height=26 class=xl1416926 style='height:22.5pt'>&nbsp;</td>                                           \n");
                htmlStr.Append("     <td class=xl746926 colspan=6>Loại tiền tệ <font                                                           \n");
                htmlStr.Append("     class='font106926'>(Currency)</font><font class='font76926'>:<span                                        \n");
                htmlStr.Append("     style='mso-spacerun:yes'> </span></font> " + dt.Rows[0]["CurrencyCodeUSD"] + "</td>                                             \n");
                htmlStr.Append("     <td class=xl776926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl746926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl656926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl746926>Tỷ giá <font                                                                           \n");
                htmlStr.Append("     class='font106926'>(Exchange rate)</font><font class='font76926'>:<span                                   \n");
                htmlStr.Append("     style='mso-spacerun:yes'> </span></font>" + dt.Rows[0]["exchangerate_no"] + "</td>                                             \n");
                htmlStr.Append("     <td class=xl656926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl656926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl656926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl656926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl656926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl656926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("     <td class=xl706926>&nbsp;</td>                                                                            \n");
                htmlStr.Append("    </tr>                                                                                                      \n");

                htmlStr.Append("     <tr height=4 style='mso-height-source:userset;height:1pt'>                                                                                                       \n");
                htmlStr.Append("      <td height=4 class=xl1416926 style='height:1pt'>&nbsp;</td>                                                                                                     \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926 width=39 style='width:36.25pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl656926 width=39 style='width:36.25pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl656926 width=67 style='width:62.5pt'>&nbsp;</td>                                                                                                      \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("      <td class=xl706926>&nbsp;</td>                                                                                                                                  \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr class=xl796926 height=21 style='height:15.75pt'>                                                                                                             \n");
                htmlStr.Append("      <td colspan=2 height=21 class=xl1716926 style='height:15.75pt'>STT</td>                                                                                         \n");
                htmlStr.Append("      <td colspan=6 class=xl1716926 style='border-right:.5pt solid black'>Tên hàng                                                                                    \n");
                htmlStr.Append("      hóa, d&#7883;ch v&#7909;</td>                                                                                                                                   \n");
                htmlStr.Append("      <td class=xl1546926>&#272;&#417;n v&#7883; tính</td>                                                                                                            \n");
                htmlStr.Append("      <td colspan=3 class=xl1716926 style='border-right:.5pt solid black'>S&#7889;                                                                                    \n");
                htmlStr.Append("      l&#432;&#7907;ng</td>                                                                                                                                           \n");
                htmlStr.Append("      <td colspan=3 class=xl1546926>&#272;&#417;n giá</td>                                                                                                            \n");
                htmlStr.Append("      <td colspan=3 class=xl1716926 style='border-right:.5pt solid black'>Thành                                                                                       \n");
                htmlStr.Append("      ti&#7873;n</td>                                                                                                                                                 \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr class=xl806926 height=17 style='height:12.75pt'>                                                                                                             \n");
                htmlStr.Append("      <td colspan=2 height=17 class=xl1736926 style='height:12.75pt'>No.</td>                                                                                         \n");
                htmlStr.Append("      <td colspan=6 class=xl1746926 style='border-right:.5pt solid black'>Description</td>                                                                            \n");
                htmlStr.Append("      <td class=xl806926>Unit</td>                                                                                                                                    \n");
                htmlStr.Append("      <td colspan=3 class=xl1736926 style='border-right:.5pt solid black'>Quantity</td>                                                                               \n");
                htmlStr.Append("      <td colspan=3 class=xl806926>Unit price</td>                                                                                                                    \n");
                htmlStr.Append("      <td colspan=3 class=xl1736926 style='border-right:.5pt solid black'>Amount</td>                                                                                 \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");
                htmlStr.Append("     <tr class=xl816926 height=20 style='mso-height-source:userset;height:15.0pt'>                                                                                    \n");
                htmlStr.Append("      <td colspan=2 height=20 class=xl1786926 style='height:15.0pt'>1</td>                                                                                            \n");
                htmlStr.Append("      <td colspan=6 class=xl1786926 style='border-right:.5pt solid black'>2</td>                                                                                      \n");
                htmlStr.Append("      <td class=xl1556926>3</td>                                                                                                                                      \n");
                htmlStr.Append("      <td colspan=3 class=xl1786926 style='border-right:.5pt solid black'>4</td>                                                                                      \n");
                htmlStr.Append("      <td colspan=3 class=xl1556926>5</td>                                                                                                                            \n");
                htmlStr.Append("      <td colspan=3 class=xl1786926 style='border-right:.5pt solid black'>6 = 4 x 5</td>                                                                              \n");
                htmlStr.Append("     </tr>                                                                                                                                                            \n");

                v_rowHeight = "29.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 26.5;

                v_rowHeightLast = "26.0pt";// "23.5pt";
                v_rowHeightLastNumber = 23.5;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "28.0pt"; //"26.5pt";    
                        v_rowHeightLast = "28.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 26.5;//27.5;
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
                        htmlStr.Append("                                                                   \n");
                        htmlStr.Append("    <tr class=xl796926 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "' >                                                               \n");
                        htmlStr.Append("      <td colspan=2 height=25 class=xl1316926 width=38 style='border-right:.5pt solid black;border-top:.5pt solid black;                                                                 \n");
                        htmlStr.Append("      height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width:36.25pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                        htmlStr.Append("      <td colspan=6 class=xl1336926 width=276 style='border-right:.5pt solid black; border-top:.5pt solid black;                                                               \n");
                        htmlStr.Append("      border-left:none;width:206pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1176926 style='border-left:none;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                        htmlStr.Append("      <td colspan=2 class=xl1806926 style='border-left:none;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][2] + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1146926 style='border-top:.5pt solid black; '>&nbsp;</td>                                                               \n");
                        htmlStr.Append("      <td colspan=2 class=xl1366926 style='border-left:none;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1136926 style='border-top:.5pt solid black; '></td>                                                               \n");
                        htmlStr.Append("      <td colspan=2 class=xl1386926 width=141 style='border-left:none;width:106pt;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1136926 style='border-top:.5pt solid black; '>&nbsp;</td>                                                               \n");
                        htmlStr.Append("     </tr>                                                               \n");
                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("    <tr class=xl746926 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                            htmlStr.Append("      <td colspan=2 height=25 class=xl1926926 width=38 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("      height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width:36.25pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                            htmlStr.Append("      <td colspan=6 class=xl1946926 width=156 style='width:116pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");

                            htmlStr.Append("                                                                     \n");
                            htmlStr.Append("      <td class=xl1206926 width=75 style='border-top:none;width:70pt'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                            //htmlStr.Append("      <td class=xl1216926 width=54 style='border-top:none;border-left:none;                                                               \n");
                            //htmlStr.Append("      width:51.25pt'></td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1256926 width=29 style='border-top:none;width:22pt'>" + dt_d.Rows[v_index][2] + "</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1226926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1236926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1156926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("     </tr>                                                               \n");
                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("    <tr class=xl746926 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                                htmlStr.Append("      <td colspan=2 height=25 class=xl1926926 width=38 style='border-right:.5pt solid black;                                                               \n");
                                htmlStr.Append("      height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width:36.25pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                                htmlStr.Append("      <td colspan=6 class=xl1946926 width=156 style='width:116pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");

                                htmlStr.Append("                                                                     \n");
                                htmlStr.Append("      <td class=xl1206926 width=75 style='border-top:none;width:70pt'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                                //htmlStr.Append("      <td class=xl1216926 width=54 style='border-top:none;border-left:none;                                                               \n");
                                //htmlStr.Append("      width:51.25pt'></td>                                                               \n");
                                htmlStr.Append("      <td colspan=2 class=xl1256926 width=29 style='border-top:none;width:22pt'>" + dt_d.Rows[v_index][2] + "</td>                                                               \n");
                                htmlStr.Append("      <td class=xl1226926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                                htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                                htmlStr.Append("      <td class=xl1236926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                                htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                                htmlStr.Append("      <td class=xl1156926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                                htmlStr.Append("     </tr>                                                               \n");
                            }
                            else
                            {
                                htmlStr.Append("    <tr class=xl746926 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                                htmlStr.Append("      <td colspan=2 height=25 class=xl1926926 width=38 style='border-right:.5pt solid black;                                                               \n");
                                htmlStr.Append("      height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width:36.25pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                                htmlStr.Append("      <td colspan=6 class=xl1946926 width=156 style='width:116pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");

                                htmlStr.Append("                                                                     \n");
                                htmlStr.Append("      <td class=xl1206926 width=75 style='border-top:none;width:70pt'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                                //htmlStr.Append("      <td class=xl1216926 width=54 style='border-top:none;border-left:none;                                                               \n");
                                //htmlStr.Append("      width:51.25pt'></td>                                                               \n");
                                htmlStr.Append("      <td colspan=2 class=xl1256926 width=29 style='border-top:none;width:22pt'>" + dt_d.Rows[v_index][2] + "</td>                                                               \n");
                                htmlStr.Append("      <td class=xl1226926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                                htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                                htmlStr.Append("      <td class=xl1236926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                                htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                                htmlStr.Append("      <td class=xl1156926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                                htmlStr.Append("     </tr>                                                               \n");
                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                        htmlStr.Append("    	<tr class=xl796926 height=25 style='mso-height-source:userset;height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
                        htmlStr.Append("      <td colspan=2 height=25 class=xl1316926 width=38 style='border-right:.5pt solid black;                                                               \n");
                        htmlStr.Append("      height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + " ;width:36.25pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                        htmlStr.Append("      <td colspan=6 class=xl1336926 width=276 style='border-right:.5pt solid black;                                                               \n");
                        htmlStr.Append("      border-left:none;width:206pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1176926 style='border-top:none;border-left:none'>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                        ///htmlStr.Append("      <td class=xl1366926 style='border-top:none;border-left:none'></td>                                                               \n");
                        htmlStr.Append("      <td colspan=2 class=xl1376926 style='border-top:none'>" + dt_d.Rows[v_index][2] + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1146926 style='border-top:none'>&nbsp;</td>                                                               \n");
                        htmlStr.Append("      <td colspan=2 class=xl1366926 style='border-left:none'>" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1136926 style='border-top:none'>&nbsp;</td>                                                               \n");
                        htmlStr.Append("      <td colspan=2 class=xl1386926 width=141 style='border-left:none;width:106pt'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                        htmlStr.Append("      <td class=xl1136926 style='border-top:none'>&nbsp;</td>                                                               \n");
                        htmlStr.Append("     </tr>                                                               \n");
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
                    v_spacePerPage = 10;  //10
                }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("    <tr class=xl746926 height=25 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                               \n");
                            htmlStr.Append("      <td colspan=2 height=25 class=xl1926926 width=38 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("      height:  " + v_rowHeightEmptyLast + "; width:36.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=4 class=xl1946926 width=156 style='width:116pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1196926 width=67 style='border-top:none;width:62.5pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1196926 width=53 style='border-top:none;width:40pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("                                                                     \n");
                            htmlStr.Append("      <td class=xl1206926 width=75 style='border-top:none;width:70pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1216926 width=54 style='border-top:none;border-left:none;                                                               \n");
                            htmlStr.Append("      width:51.25pt'></td>                                                               \n");
                            htmlStr.Append("      <td class=xl1256926 width=29 style='border-top:none;width:22pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1226926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1236926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1156926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("     </tr>                                                               \n");
                        }
                        else
                        {
                            htmlStr.Append("    	<tr class=xl796926 height=25 style='mso-height-source:userset;height: " + v_rowHeightEmptyLast + "'>                                                               \n");
                            htmlStr.Append("      <td colspan=2 height=25 class=xl1316926 width=38 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("      height:" + v_rowHeightEmptyLast + " ;width:36.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=6 class=xl1336926 width=276 style='border-right:.5pt solid black;                                                               \n");
                            htmlStr.Append("      border-left:none;width:206pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1176926 style='border-top:none;border-left:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1366926 style='border-top:none;border-left:none'></td>                                                               \n");
                            htmlStr.Append("      <td class=xl1376926 style='border-top:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1146926 style='border-top:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1366926 style='border-left:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1136926 style='border-top:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1386926 width=141 style='border-left:none;width:106pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1136926 style='border-top:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("     </tr>                                                               \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    htmlStr.Append("    <tr class=xl746926 height=26 style='mso-height-source:userset;height:" + (v_spacePerPage).ToString() + "pt'>                                                               \n");
                    htmlStr.Append("      <td height=26 class=xl976926 width=6 style='height:" + (v_spacePerPage).ToString() + "pt;width:6.25pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td class=xl1246926 width=32 style='width:30pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td class=xl1246926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td class=xl1246926 width=53 style='width:40pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td class=xl1246926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td class=xl1246926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td class=xl1246926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td class=xl1246926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
                    htmlStr.Append("      <td colspan=7 class=xl1876926 width=255 style='                                                               \n");
                    htmlStr.Append("      width:193pt'></td>                                                               \n");
                    htmlStr.Append("      <td colspan=2 class=xl1896926 style='border-left:none'></td>                                                               \n");
                    htmlStr.Append("      <td class=xl1166926>&nbsp;</td>                                                               \n");
                    htmlStr.Append("     </tr>                                                       \n");

                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=18 style='height: 10pt'>                                                                                                                                                                \n");
                    htmlStr.Append("			<td colspan=27 height=18                                                                                                                                                       \n");
                    htmlStr.Append("				style=' height: 10pt'>&nbsp;</td>                                                                                                                           \n");
                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             

            htmlStr.Append("    <tr class=xl746926 height=26 style='mso-height-source:userset;height:22.1pt'>                                                               \n");
            htmlStr.Append("      <td height=26 class=xl976926 width=6 style='height:24.1pt;width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1246926 width=32 style='width:30pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1246926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1246926 width=53 style='width:40pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1246926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1246926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1246926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1246926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl1876926 width=255 style='border-right:.5pt solid black;                                                               \n");
            htmlStr.Append("      width:193pt'>C&#7897;ng ti&#7873;n hàng (<font class='font106926'>Total                                                               \n");
            htmlStr.Append("      amount</font><font class='font86926'>) :</font></td>                                                               \n");
            htmlStr.Append("      <td colspan=2 class=xl1896926 style='border-left:none'> " + dt.Rows[0]["netamount_display"] + "&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1166926>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl746926 height=26 style='mso-height-source:userset;height:22.1pt'>                                                               \n");
            htmlStr.Append("      <td height=26 class=xl976926 width=6 style='height:24.1pt;border-top:none;                                                               \n");
            htmlStr.Append("      width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=4 class=xl1306926 width=191 style='width:143pt'><span                                                               \n");
            htmlStr.Append("      style='mso-spacerun:yes'> </span>Thu&#7871; su&#7845;t GTGT (<font                                                               \n");
            htmlStr.Append("      class='font106926'>VAT rate</font><font class='font86926'>) :</font></td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=39 style='border-top:none;width:36.25pt'>" + dt.Rows[0]["TaxRate"] + "&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1266926 width=67 style='border-top:none;width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=11 style='border-top:none;width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl1876926 width=255 style='border-right:.5pt solid black;                                                               \n");
            htmlStr.Append("      width:193pt'>Ti&#7873;n thu&#7871; GTGT (<font class='font106926'>VAT</font><font                                                               \n");
            htmlStr.Append("      class='font86926'>) :</font></td>                                                               \n");
            htmlStr.Append("      <td colspan=2 class=xl1896926 style='border-left:none'>" + dt.Rows[0]["vatamount_display"] + "&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl836926 style='border-top:none'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl746926 height=26 style='mso-height-source:userset;height:22.1pt'>                                                               \n");
            htmlStr.Append("      <td height=26 class=xl976926 width=6 style='height:24.1pt;border-top:none;                                                               \n");
            htmlStr.Append("      width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=32 style='border-top:none;width:30pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=67 style='border-top:none;width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=53 style='border-top:none;width:40pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=39 style='border-top:none;width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=39 style='border-top:none;width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=67 style='border-top:none;width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1306926 width=11 style='border-top:none;width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl1876926 width=255 style='border-right:.5pt solid black;                                                               \n");
            htmlStr.Append("      width:193pt'>T&#7893;ng ti&#7873;n thanh toán (<font class='font106926'>Total                                                               \n");
            htmlStr.Append("      payment</font><font class='font86926'>) :</font></td>                                                               \n");
            htmlStr.Append("      <td colspan=2 class=xl1896926 style='border-left:none'>" + dt.Rows[0]["totalamount_display"] + "&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl836926>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl746926 height=24 style='mso-height-source:userset;height:22.0pt'>                                                               \n");
            htmlStr.Append("      <td height=24 class=xl826926 style='height:24.0pt;border-top:none'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=16 class=xl1916926>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng                                                               \n");
            htmlStr.Append("      ch&#7919; (<font class='font106926'>In words</font><font class='font76926'>):                                                               \n");
            htmlStr.Append("      </font>&nbsp;" + read_prive + "</td>                                                               \n");
            htmlStr.Append("      <td class=xl846926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr height=18 style='mso-height-source:userset;height:24.1pt'>                                                               \n");
            htmlStr.Append("      <td height=18 class=xl786926 style='height:24.1pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl856926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl856926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl856926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=75 style='width:70pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=54 style='width:51.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=29 style='width:22pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=46 style='width:35pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl866926 width=102 style='width:96.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl876926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl766926 height=20 style='mso-height-source:userset;height:15.6pt'>                                                               \n");
            htmlStr.Append("      <td height=20 class=xl896926 style='height:15.6pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=5 class=xl2046926 width=230 style='width:172pt'>Ng&#432;&#7901;i                                                               \n");
            htmlStr.Append("      mua hàng (<font class='font146926'>Buyer</font><font class='font56926'>)</font></td>                                                               \n");
            htmlStr.Append("      <td colspan=4 class=xl2006926 width=207 style='width:155pt'></td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl2006926 width=267 style='width:202pt'>Ng&#432;&#7901;i                                                               \n");
            htmlStr.Append("      bán hàng (<font class='font146926'>Seller</font><font class='font56926'>)</font></td>                                                               \n");
            htmlStr.Append("      <td class=xl886926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl1526926 height=18 style='mso-height-source:userset;height:14.1pt'>                                                               \n");
            htmlStr.Append("      <td height=18 class=xl1516926 style='height:14.1pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=5 class=xl1856926 width=230 style='width:172pt'>(Ký, ghi rõ                                                               \n");
            htmlStr.Append("      h&#7885; tên)</td>                                                               \n");
            htmlStr.Append("      <td colspan=4 class=xl1856926 width=207 style='width:155pt'></td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl1856926 width=267 style='width:202pt'>(Ký, &#273;óng                                                               \n");
            htmlStr.Append("      d&#7845;u, ghi rõ h&#7885; tên)</td>                                                               \n");
            htmlStr.Append("      <td class=xl906926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr class=xl936926 height=20 style='mso-height-source:userset;height:15.6pt'>                                                               \n");
            htmlStr.Append("      <td height=20 class=xl916926 style='height:15.6pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=5 class=xl1866926 width=230 style='width:172pt'>(Signature &amp;                                                               \n");
            htmlStr.Append("      full name)</td>                                                               \n");
            htmlStr.Append("      <td colspan=4 class=xl1866926 width=207 style='width:155pt'></td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl1866926 width=267 style='width:202pt'>(Signature, stamp                                                               \n");
            htmlStr.Append("      &amp; full name)</td>                                                               \n");
            htmlStr.Append("      <td class=xl926926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr height=40 style='mso-height-source:userset;height:30.0pt'>                                                               \n");
            htmlStr.Append("      <td height=40 class=xl1416926 style='height:30.0pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=32 style='width:30pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=53 style='width:40pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=75 style='width:70pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=54 style='width:51.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=29 style='width:22pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=46 style='width:35pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=102 style='width:96.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl946926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("                                                                    \n");


            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {
                htmlStr.Append("    					 <tr height=22 style='height:16.5pt'>                                                               \n");
                htmlStr.Append("      <td height=22 class=xl1416926 style='height:16.5pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=75 style='width:70pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1296926 colspan=3>Signature Valid</td>                                                               \n");
                htmlStr.Append("      <td align=left valign=top class=xl1006926><![if !vml]><span style='mso-ignore:vglayout;                                                               \n");
                htmlStr.Append("      position:absolute;z-index:1;margin-left:21px;margin-top:17px;width:55px;                                                               \n");
                htmlStr.Append("      height:40px'><img width=55 height=40                                                               \n");
                htmlStr.Append("          src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png'                                                              \n");
                htmlStr.Append("      v:shapes='Picture_x0020_8'></span><![endif]><span style='mso-ignore:vglayout2'>                                                               \n");
                htmlStr.Append("      <table cellpadding=0 cellspacing=0>                                                               \n");
                htmlStr.Append("       <tr>                                                               \n");
                htmlStr.Append("        <td height=22  width=46 style='height:16.5pt;width:35pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("       </tr>                                                               \n");
                htmlStr.Append("      </table>                                                               \n");
                htmlStr.Append("      </span></td>                                                               \n");
                htmlStr.Append("      <td class=xl1016926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1016926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1026926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1046926>&nbsp;</td>                                                               \n");
                htmlStr.Append("     </tr>                                                               \n");
            }
            else
            {
                htmlStr.Append("    					 <tr height=22 style='height:16.5pt'>                                                               \n");
                htmlStr.Append("      <td height=22 class=xl1416926 style='height:16.5pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl716926 width=75 style='width:70pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1296926 colspan=3>Signature Valid</td>                                                               \n");
                htmlStr.Append("      <td align=left valign=top class=xl1006926><![if !vml]><span style='mso-ignore:vglayout;                                                               \n");
                htmlStr.Append("      position:absolute;z-index:1;margin-left:21px;margin-top:17px;width:55px;                                                               \n");
                htmlStr.Append("      height:40px'></span><![endif]><span style='mso-ignore:vglayout2'>                                                               \n");
                htmlStr.Append("      <table cellpadding=0 cellspacing=0>                                                               \n");
                htmlStr.Append("       <tr>                                                               \n");
                htmlStr.Append("        <td height=22  width=46 style='height:16.5pt;width:35pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("       </tr>                                                               \n");
                htmlStr.Append("      </table>                                                               \n");
                htmlStr.Append("      </span></td>                                                               \n");
                htmlStr.Append("      <td class=xl1016926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1016926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1026926>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl1046926>&nbsp;</td>                                                               \n");
                htmlStr.Append("     </tr>                                                               \n");
            }

            htmlStr.Append("                                                                   \n");
            htmlStr.Append("     <tr class=xl1086926 height=21 style='height:15.75pt'>                                                               \n");
            htmlStr.Append("      <td height=21 class=xl1076926 style='height:15.75pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1086926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1086926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1086926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1096926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1096926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1096926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1096926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1096926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1086926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=7 class=xl2016926 style='border-right:.5pt solid black'>&#272;&#432;&#7907;c                                                               \n");
            htmlStr.Append("      ký b&#7903;i: <font class='font126926'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                               \n");
            htmlStr.Append("      <td class=xl1106926>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr height=21 style='height:15.75pt'>                                                               \n");
            htmlStr.Append("      <td height=21 class=xl1416926 style='height:15.75pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("       <td class=xl13821177></td>                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=75 style='width:70pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1286926 colspan=3>Ngày ký: " + dt.Rows[0]["SignedDate"] + "<span style='mso-spacerun:yes'> </span></td>                                                               \n");
            htmlStr.Append("      <td class=xl1496926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1496926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1496926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1506926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1036926>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("                                                                   \n");
            htmlStr.Append("     <tr height=20 style='mso-height-source:userset;height:15.0pt'>                                                               \n");
            htmlStr.Append("      <td height=20 class=xl1416926 style='height:15.0pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl6569261 colspan=8>Tra c&#7913;u t&#7841;i Website: <font                                                               \n");
            htmlStr.Append("      class='font66926'><span style='mso-spacerun:yes'> </span></font><font                                                               \n");
            htmlStr.Append("      class='font116926'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></td>                                                               \n");
            htmlStr.Append("      <td class=xl656926>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl1066926 colspan=4>Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl716926 width=102 style='width:96.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl946926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr height=16 style='mso-height-source:userset;height:12.0pt'>                                                               \n");
            htmlStr.Append("      <td height=16 class=xl786926 style='height:12.0pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=16 class=xl1976926>(C&#7847;n ki&#7875;m tra, &#273;&#7889;i                                                               \n");
            htmlStr.Append("      chi&#7871;u khi l&#7853;p, giao nh&#7853;n hóa &#273;&#417;n)</td>                                                               \n");
            htmlStr.Append("      <td class=xl966926 width=6 style='width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <tr height=16 style='mso-height-source:userset;height:12.0pt'>                                                               \n");
            htmlStr.Append("      <td height=16 class=xl676926 style='height:12.0pt;border-top:none'>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td colspan=16 class=xl1986926>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                               \n");
            htmlStr.Append("      <td class=xl1276926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <![if supportMisalignedColumns]>                                                               \n");
            htmlStr.Append("     <tr height=0 style='display:none'>                                                               \n");
            htmlStr.Append("      <td width=6 style='width:6.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=32 style='width:30pt'></td>                                                               \n");
            htmlStr.Append("      <td width=67 style='width:62.5pt'></td>                                                               \n");
            htmlStr.Append("      <td width=53 style='width:40pt'></td>                                                               \n");
            htmlStr.Append("      <td width=39 style='width:36.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=39 style='width:36.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=67 style='width:62.5pt'></td>                                                               \n");
            htmlStr.Append("      <td width=11 style='width:10pt'></td>                                                               \n");
            htmlStr.Append("      <td width=75 style='width:70pt'></td>                                                               \n");
            htmlStr.Append("      <td width=54 style='width:51.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=29 style='width:22pt'></td>                                                               \n");
            htmlStr.Append("      <td width=6 style='width:6.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=39 style='width:36.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=46 style='width:35pt'></td>                                                               \n");
            htmlStr.Append("      <td width=6 style='width:6.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=39 style='width:36.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=102 style='width:96.25pt'></td>                                                               \n");
            htmlStr.Append("      <td width=6 style='width:6.25pt'></td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("     <![endif]>                                                               \n");
            htmlStr.Append("    </table>                                                               \n");
            htmlStr.Append("   </ body >          \n");
            htmlStr.Append("    </ html >               \n");

            string filePath = "";

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\webproject\e-invoice-ws\02.Web\AttachFileText\" + tei_einvoice_m_pk + ".html"))
            {
                file.WriteLine(htmlStr.ToString()); // "sb" is the StringBuilder
            }

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