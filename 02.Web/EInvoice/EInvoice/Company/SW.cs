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
    public class SW
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


            string read_prive = "", read_en = "", read_amount = "", amount_vat = "", amount_total = "", amount_trans = "", amount_net = "", lb_amount_trans = "";


            if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            {
                lb_amount_trans = "";
                amount_trans = "";
                amount_total = dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                amount_vat = dt.Rows[0]["VAT_BK_AMT_92"].ToString();
                amount_net = dt.Rows[0]["NET_BK_AMT_90"].ToString();
                // read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmountInWord"].ToString()));
            }
            else
            {
                lb_amount_trans = "Tổng cộng VND <font class='font1127974'>(Amount VND):</font>";
                amount_trans = dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                amount_total = dt.Rows[0]["tot_amt_tr_94"].ToString();
                amount_vat = dt.Rows[0]["VAT_TR_AMT_DIS_TR_91"].ToString();
                amount_net = dt.Rows[0]["NET_TR_AMT_DIS_TR_89"].ToString();

                // read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            }
            read_prive = dt.Rows[0]["AMOUNT_WORD_VIE"].ToString();//;

            //read_prive = dt.Rows[0]["amount_word_vie"].ToString();

            //read_en = dt.Rows[0]["amount_word_eng"].ToString();

            /*if (dt.Rows[0]["vatamount_display"].ToString().Trim() == "0")
            {
                amout_vat = "-";
            }
            else
            {
                amout_vat = dt.Rows[0]["vatamount_display"].ToString();
            }*/

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

            htmlStr.Append("      <!--table																																			\n");
            htmlStr.Append("     	{mso-displayed-decimal-separator:'\\.';                                                                                                         \n");
            htmlStr.Append("     	mso-displayed-thousand-separator:'\\,';}                                                                                                        \n");
            htmlStr.Append("     .font51637                                                                                                                                         \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font61637                                                                                                                                         \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font71637                                                                                                                                         \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font81637                                                                                                                                         \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font91637                                                                                                                                         \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font101637                                                                                                                                        \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font111637                                                                                                                                        \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font121637                                                                                                                                        \n");
            htmlStr.Append("     	{color:#0066CC;                                                                                                                                 \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font131637                                                                                                                                        \n");
            htmlStr.Append("     	{color:red;                                                                                                                                     \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font141637                                                                                                                                        \n");
            htmlStr.Append("     	{color:windowtext;                                                                                                                              \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font151637                                                                                                                                        \n");
            htmlStr.Append("     	{color:black;                                                                                                                                   \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font161637                                                                                                                                        \n");
            htmlStr.Append("     	{color:red;                                                                                                                                     \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .font171637                                                                                                                                        \n");
            htmlStr.Append("     	{color:black;                                                                                                                                   \n");
            htmlStr.Append("     	font-size:10.5pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;}                                                                                                                            \n");
            htmlStr.Append("     .xl151637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial;                                                                                                                              \n");
            htmlStr.Append("     	mso-generic-font-family:auto;                                                                                                                   \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl651637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl661637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl671637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl681637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl691637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl701637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl711637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl721637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl731637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl741637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl751637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl761637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl771637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl781637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl791637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl801637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl811637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:#C00000;                                                                                                                                  \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl821637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl831637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl841637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl851637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl861637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl871637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl881637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl891637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl901637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl911637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl921637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:#0070C0;                                                                                                                                  \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl931637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl941637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl951637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl961637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl971637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl981637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl991637                                                                                                                                          \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1001637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1011637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1021637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1031637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1041637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1051637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1061637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1071637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1081637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1091637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:0%;                                                                                                                           \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1101637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1111637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:black;                                                                                                                                    \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1121637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:black;                                                                                                                                    \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:bottom;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1131637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:#002060;                                                                                                                                  \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial, sans-serif;                                                                                                                  \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1141637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1151637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'Short Date';                                                                                                                 \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1161637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'Short Date';                                                                                                                 \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1171637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1181637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1191637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1201637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1211637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1221637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1231637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1241637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1251637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1261637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.00pt;                                                                                                                              \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1271637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1281637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1291637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1301637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1311637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:18.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1321637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1331637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1341637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:#ED7D31;                                                                                                                                  \n");
            htmlStr.Append("     	font-size:13.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1351637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:black;                                                                                                                                    \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1361637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:black;                                                                                                                                    \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1371637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:black;                                                                                                                                    \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1381637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1391637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1401637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1411637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1421637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1431637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1441637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1451637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1461637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1471637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1481637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1491637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1501637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1511637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1521637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1531637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.50pt;                                                                                                                              \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1541637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1551637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1561637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1571637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1581637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\#\\,\\#\\#0';                                                                                                              \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1591637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\#\\,\\#\\#0';                                                                                                              \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1601637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial;                                                                                                                              \n");
            htmlStr.Append("     	mso-generic-font-family:auto;                                                                                                                   \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1611637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial;                                                                                                                              \n");
            htmlStr.Append("     	mso-generic-font-family:auto;                                                                                                                   \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1621637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial;                                                                                                                              \n");
            htmlStr.Append("     	mso-generic-font-family:auto;                                                                                                                   \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1631637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\#\\,\\#\\#0';                                                                                                              \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1641637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1651637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1661637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1671637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1681637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1691637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1701637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:1.0pt dotted windowtext;                                                                                                           \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1711637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1721637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1731637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1741637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1751637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1761637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1771637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1781637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1791637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1801637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1811637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                                \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1821637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                                \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1831637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                                \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1841637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:17.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1851637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:red;                                                                                                                                      \n");
            htmlStr.Append("     	font-size:14.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1861637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:italic;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1871637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:12.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1881637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:black;                                                                                                                                    \n");
            htmlStr.Append("     	font-size:10.5pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1891637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1901637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial;                                                                                                                              \n");
            htmlStr.Append("     	mso-generic-font-family:auto;                                                                                                                   \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1911637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial;                                                                                                                              \n");
            htmlStr.Append("     	mso-generic-font-family:auto;                                                                                                                   \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1921637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:'\\@';                                                                                                                        \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1931637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:Arial;                                                                                                                              \n");
            htmlStr.Append("     	mso-generic-font-family:auto;                                                                                                                   \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:right;                                                                                                                               \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:1.0pt dotted windowtext;                                                                                                              \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:.5pt solid windowtext;                                                                                                            \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	mso-background-source:auto;                                                                                                                     \n");
            htmlStr.Append("     	mso-pattern:auto;                                                                                                                               \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl1941637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                                \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:.5pt solid windowtext;                                                                                                              \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1951637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                                \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:none;                                                                                                                              \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1961637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:1.0pt;                                                                                                                                \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:.5pt solid windowtext;                                                                                                               \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1971637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:normal;}                                                                                                                            \n");
            htmlStr.Append("     .xl1981637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl1991637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:general;                                                                                                                             \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	border-top:none;                                                                                                                                \n");
            htmlStr.Append("     	border-right:.5pt solid windowtext;                                                                                                             \n");
            htmlStr.Append("     	border-bottom:none;                                                                                                                             \n");
            htmlStr.Append("     	border-left:none;                                                                                                                               \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl2001637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:center;                                                                                                                              \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;                                                                                                                             \n");
            htmlStr.Append("     	mso-text-control:shrinktofit;}                                                                                                                  \n");
            htmlStr.Append("     .xl2011637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:11.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:700;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     .xl2021637                                                                                                                                         \n");
            htmlStr.Append("     	{padding:0px;                                                                                                                                   \n");
            htmlStr.Append("     	mso-ignore:padding;                                                                                                                             \n");
            htmlStr.Append("     	color:windowtext;                                                                                                                               \n");
            htmlStr.Append("     	font-size:10.0pt;                                                                                                                               \n");
            htmlStr.Append("     	font-weight:400;                                                                                                                                \n");
            htmlStr.Append("     	font-style:normal;                                                                                                                              \n");
            htmlStr.Append("     	text-decoration:none;                                                                                                                           \n");
            htmlStr.Append("     	font-family:'Times New Roman', serif;                                                                                                           \n");
            htmlStr.Append("     	mso-font-charset:0;                                                                                                                             \n");
            htmlStr.Append("     	mso-number-format:General;                                                                                                                      \n");
            htmlStr.Append("     	text-align:left;                                                                                                                                \n");
            htmlStr.Append("     	vertical-align:middle;                                                                                                                          \n");
            htmlStr.Append("     	background:white;                                                                                                                               \n");
            htmlStr.Append("     	mso-pattern:black none;                                                                                                                         \n");
            htmlStr.Append("     	white-space:nowrap;}                                                                                                                            \n");
            htmlStr.Append("     -->                                                                                                                                                \n"); 
            htmlStr.Append("    </style>                                                               \n");
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

                htmlStr.Append("      	<table border=0 cellpadding=0 cellspacing=0 width=732 style='border-collapse:																															\n");
                htmlStr.Append("      	 collapse;table-layout:fixed;width:538pt'>                                                                                                                                                              \n");
                htmlStr.Append("      	 <col class=xl651637 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                             \n");
                htmlStr.Append("      	 199;width:4pt'>                                                                                                                                                                                        \n");
                htmlStr.Append("      	 <col class=xl651637 width=33 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1166;width:25pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=70 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 2474;width:52pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=55 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1962;width:41pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=34 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1223;width:26pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                             \n");
                htmlStr.Append("      	 227;width:5pt'>                                                                                                                                                                                        \n");
                htmlStr.Append("      	 <col class=xl651637 width=41 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1450;width:31pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=70 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 2474;width:52pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=12 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 426;width:9pt'>                                                                                                                                                                                        \n");
                htmlStr.Append("      	 <col class=xl651637 width=78 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 2759;width:63pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=56 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1991;width:32pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=30 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1080;width:23pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                             \n");
                htmlStr.Append("      	 199;width:4pt'>                                                                                                                                                                                        \n");
                htmlStr.Append("      	 <col class=xl651637 width=41 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1450;width:31pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=49 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1735;width:37pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                             \n");
                htmlStr.Append("      	 199;width:4pt'>                                                                                                                                                                                        \n");
                htmlStr.Append("      	 <col class=xl651637 width=41 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 1450;width:31pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=92 style='mso-width-source:userset;mso-width-alt:                                                                                                                            \n");
                htmlStr.Append("      	 3271;width:65pt'>                                                                                                                                                                                      \n");
                htmlStr.Append("      	 <col class=xl651637 width=6 style='mso-width-source:userset;mso-width-alt:                                                                                                                             \n");
                htmlStr.Append("      	 199;width:4pt'>                                                                                                                                                                                        \n");
                htmlStr.Append("      	 <tr height=30 style='height:22.8pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=30 class=xl971637 width=6 style='height:22.8pt;width:4pt'>&nbsp;</td>                                                                                                                      \n");
                htmlStr.Append("      	  <td width=33 style='width:25pt' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                                                   \n");
                htmlStr.Append("      	  position:absolute;z-index:2;margin-left:2px;margin-top:8px;width:109px;                                                                                                                               \n");
                htmlStr.Append("      	  height:73px'><img width=109 height=73                                                                                                                                                                 \n");
                htmlStr.Append("      	  src='${pageContext.request.contextPath}/assets/images/Bornga_001.png'                                                                                                                                 \n");
                htmlStr.Append("      	  v:shapes='Picture_x0020_2'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                                                                       \n");
                htmlStr.Append("      	  <table cellpadding=0 cellspacing=0>                                                                                                                                                                   \n");
                htmlStr.Append("      	   <tr>                                                                                                                                                                                                 \n");
                htmlStr.Append("      		<td height=30 class=xl1201637 width=33 style='height:22.8pt;width:25pt'>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("      	   </tr>                                                                                                                                                                                                \n");
                htmlStr.Append("      	  </table>                                                                                                                                                                                              \n");
                htmlStr.Append("      	  </span></td>                                                                                                                                                                                          \n");
                htmlStr.Append("      	  <td class=xl841637 width=70 style='width:52pt'>&nbsp;</td>                                                                                                                                            \n");
                htmlStr.Append("      	  <td class=xl1201637 width=55 style='width:41pt'>&nbsp;</td>                                                                                                                                           \n");
                htmlStr.Append("      	  <td colspan=10 class=xl1841637 width=374 style='width:281pt'>HÓA                                                                                                                                      \n");
                htmlStr.Append("      	  &#272;&#416;N GIÁ TR&#7882; GIA T&#258;NG</td>                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl1321637 width=49 style='width:37pt'>&nbsp;</td>                                                                                                                                           \n");
                htmlStr.Append("      	  <td class=xl1201637 width=6 style='width:4pt'>&nbsp;</td>                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1201637 width=41 style='width:31pt'>&nbsp;</td>                                                                                                                                           \n");
                htmlStr.Append("      	  <td class=xl1331637 width=92 style='width:69pt'>&nbsp;</td>                                                                                                                                           \n");
                htmlStr.Append("      	  <td class=xl851637 width=6 style='width:4pt'>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=24 style='height:18.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=24 class=xl1191637 style='height:18.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=10 class=xl1851637>(VAT INVOICE)</td>                                                                                                                                                     \n");
                htmlStr.Append("      	  <td class=xl661637 colspan=3>Ký hi&#7879;u (<font class='font111637'>Serial)</font></td>                                                                                                              \n");
                htmlStr.Append("      	  <td class=xl1181637>: " + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["INVOICESERIALNO"] + "</td>                                                                                                                                            \n");
                htmlStr.Append("      	  <td class=xl671637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=11 class=xl1861637>&nbsp;</td>                                                                                                                                                            \n");
                htmlStr.Append("      	  <td class=xl661637 colspan=2>S&#7889; <font class='font111637'>(No</font><font                                                                                                                        \n");
                htmlStr.Append("      	  class='font81637'>)</font></td>                                                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl1141637><font class='font151637'>:</font><font class='font131637'><span                                                                                                                   \n");
                htmlStr.Append("      	  style='mso-spacerun:yes'> </span>" + dt.Rows[0]["InvoiceNumber"] + "</font></td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl671637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=10 class=xl1871637 width=374 style='width:281pt'>Ngày <font                                                                                                                               \n");
                htmlStr.Append("      	  class='font111637'>(Date)</font><font class='font71637'><span                                                                                                                                         \n");
                htmlStr.Append("      	  style='mso-spacerun:yes'>  " + dt.Rows[0]["INVOICEISSUEDDATE_DD"] + " </span></font><font class='font71637'><span                                                                                                                  \n");
                htmlStr.Append("      	  style='mso-spacerun:yes'> </span>tháng </font><font class='font111637'>(month)</font><font                                                                                                            \n");
                htmlStr.Append("      	  class='font71637'><span style='mso-spacerun:yes'>  " + dt.Rows[0]["INVOICEISSUEDDATE_MM"] + "  </span>n&#259;m </font><font                                                                                                        \n");
                htmlStr.Append("      	  class='font111637'>(year)</font><font class='font71637'><span                                                                                                                                         \n");
                htmlStr.Append("      	  style='mso-spacerun:yes'>  " + dt.Rows[0]["INVOICEISSUEDDATE_YYYY"] + "  </span></font></td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl691637 width=6 style='width:4pt'>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=18 style='height:13.8pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td colspan=2 height=18 class=xl1191637 style='height:13.8pt'>&nbsp;</td>                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=10 class=xl1881637>MCCQT: <font class='font171637'>" + dt.Rows[0]["cqt_mccqt_id"] + "</font></td>                                                                                                               \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;" + v_titlePageNumber + "</td>                                                                                                                                                 \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl671637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                           \n");
                htmlStr.Append("      	  <td colspan=19 height=4 width=732 style='border-right:.5pt solid black;                                                                                                                               \n");
                htmlStr.Append("      	  height:3.0pt;width:535pt' align=left valign=top><![if !vml]><span style='mso-ignore:vglayout;                                                                                                         \n");
                htmlStr.Append("      	  position:absolute;z-index:3;margin-left:2px;margin-top:3px;width:715px;                                                                                                                               \n");
                htmlStr.Append("      	  height:715px'><img width=715 height=715                                                                                                                                                               \n");
                htmlStr.Append("      	  src='${pageContext.request.contextPath}/assets/images/Bornga_002.png'                                                                                                                                 \n");
                htmlStr.Append("      	  alt=''                                                                                                                                                                                                \n");
                htmlStr.Append("      	  v:shapes='AutoShape_x0020_85 _x0000_s14573 Picture_x0020_5'></span><![endif]><span                                                                                                                    \n");
                htmlStr.Append("      	  style='mso-ignore:vglayout2'>                                                                                                                                                                         \n");
                htmlStr.Append("      	  <table cellpadding=0 cellspacing=0>                                                                                                                                                                   \n");
                htmlStr.Append("      	   <tr>                                                                                                                                                                                                 \n");
                htmlStr.Append("      		<td colspan=19 height=4 class=xl1811637 width=732 style='border-right:.5pt solid black;                                                                                                             \n");
                htmlStr.Append("      		height:3.0pt;width:538pt'>&nbsp;</td>                                                                                                                                                               \n");
                htmlStr.Append("      	   </tr>                                                                                                                                                                                                \n");
                htmlStr.Append("      	  </table>                                                                                                                                                                                              \n");
                htmlStr.Append("      	  </span></td>                                                                                                                                                                                          \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                           \n");
                htmlStr.Append("      	  <td colspan=19 height=4 class=xl1941637 style='border-right:.5pt solid black;                                                                                                                         \n");
                htmlStr.Append("      	  height:3.0pt'>&nbsp;</td>                                                                                                                                                                             \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=28 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("      	  <td height=28 class=xl1291637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=4>Tên &#273;&#417;n v&#7883; bán hàng <font                                                                                                                                \n");
                htmlStr.Append("      	  class='font101637'>(Seller)</font><font class='font71637'>:</font></td>                                                                                                                               \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl1341637>" + dt.Rows[0]["Seller_Name"] + "</td>                                                                                                                                                           \n");
                htmlStr.Append("      	  <td class=xl1131637></td>                                                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1131637></td>                                                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1131637></td>                                                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1131637></td>                                                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1131637></td>                                                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1131637></td>                                                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1131637></td>                                                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1301637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl1301637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl1301637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl1301637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl1311637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=3>Mã s&#7889; thu&#7871; <font class='font101637'>(Tax                                                                                                                     \n");
                htmlStr.Append("      	  code) :</font></td>                                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=12 class=xl1781637 width=522 style='width:391pt'>" + dt.Rows[0]["Seller_Taxcode"] + "</td>                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl671637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=33 style='mso-height-source:userset;height:25.05pt'>                                                                                                                                        \n");
                htmlStr.Append("      	  <td height=33 class=xl1191637 style='height:25.05pt'>&nbsp;</td>                                                                                                                                      \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=2>&#272;&#7883;a ch&#7881; <font class='font101637'>(Address) :</font></td>                                                                                                \n");
                htmlStr.Append("      	  <td class=xl151637></td>                                                                                                                                                                              \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=12 class=xl1791637 width=522 style='width:391pt'>" + dt.Rows[0]["SELLER_ADDRESS"] + "</td>                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl671637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=3>&#272;i&#7879;n tho&#7841;i <font                                                                                                                                        \n");
                htmlStr.Append("      	  class='font101637'>(Tel) :</font></td>                                                                                                                                                                \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=12 class=xl1801637 width=522 style='width:391pt'>" + dt.Rows[0]["Seller_Tel"] + "</td>                                                                                                                 \n");
                htmlStr.Append("      	  <td class=xl691637 width=6 style='width:4pt'>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=4>S&#7889; tài kho&#7843;n <font class='font101637'>(Acc.                                                                                                                  \n");
                htmlStr.Append("      	  code) :</font></td>                                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=12 class=xl1801637 width=522 style='width:391pt'>" + dt.Rows[0]["SELLER_ACCOUNTNO"] + "</td>                                                                                                                 \n");
                htmlStr.Append("      	  <td class=xl691637 width=6 style='width:4pt'>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl651637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl701637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td class=xl701637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	  <td colspan=12 class=xl1801637 width=522 style='width:391pt'>" + dt.Rows[0]["BANK_NM78"] + "</td>                                                                                                                    \n");
                htmlStr.Append("      	  <td class=xl691637 width=6 style='width:4pt'>&nbsp;</td>                                                                                                                                              \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                           \n");
                htmlStr.Append("      	  <td colspan=19 height=4 class=xl1811637 style='border-right:.5pt solid black;                                                                                                                         \n");
                htmlStr.Append("      	  height:3.0pt'>&nbsp;</td>                                                                                                                                                                             \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                           \n");
                htmlStr.Append("      	  <td colspan=19 height=4 class=xl1941637 style='border-right:.5pt solid black;                                                                                                                         \n");
                htmlStr.Append("      	  height:3.0pt'>&nbsp;</td>                                                                                                                                                                             \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=7>H&#7885; tên ng&#432;&#7901;i mua hàng<font                                                                                                                              \n");
                htmlStr.Append("      	  class='font101637'> (Customer's name)</font><font class='font71637'>:</font></td>                                                                                                                     \n");
                htmlStr.Append("      	  <td colspan=10 class=xl2001637>&nbsp;" + dt.Rows[0]["buyer"] + "</td>                                                                                                                                          \n");
                htmlStr.Append("      	  <td class=xl1991637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=5>Tên &#273;&#417;n v&#7883; <font                                                                                                                                         \n");
                htmlStr.Append("      	  class='font101637'>(Company's name)</font><font class='font71637'>:</font></td>                                                                                                                       \n");
                htmlStr.Append("      	  <td colspan=12 class=xl1971637>&nbsp;" + dt.Rows[0]["buyerlegalname"] + "</td>                                                                                                                                           \n");
                htmlStr.Append("      	  <td class=xl1981637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=3>Mã s&#7889; thu&#7871; <font class='font101637'>(Tax                                                                                                                     \n");
                htmlStr.Append("      	  code)</font><font class='font71637'>:<span style='mso-spacerun:yes'> </span></font></td>                                                                                                              \n");
                htmlStr.Append("      	  <td colspan=14 class=xl1971637>&nbsp;" + dt.Rows[0]["BuyerTaxCode"] + "</td>                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl1981637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=3>&#272;&#7883;a ch&#7881;<font class='font101637'>                                                                                                                        \n");
                htmlStr.Append("      	  (Address)</font><font class='font71637'>:<span                                                                                                                                                        \n");
                htmlStr.Append("      	  style='mso-spacerun:yes'> </span></font></td>                                                                                                                                                         \n");
                htmlStr.Append("      	  <td colspan=14 class=xl1971637>&nbsp;" + dt.Rows[0]["BuyerAddress"] + "</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl1211637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=7>C&#259;n c&#432;&#7899;c công dân <font                                                                                                                                  \n");
                htmlStr.Append("      	  class='font111637'>(Citizen identification)</font><font class='font71637'>:</font></td>                                                                                                               \n");
                htmlStr.Append("      	  <td colspan=10 class=xl1971637>&nbsp;</td>                                                                                                                                                            \n");
                htmlStr.Append("      	  <td class=xl1211637>&nbsp;</td>                                                                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=22 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td height=22 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=7>Hình th&#7913;c thanh toán <font                                                                                                                                         \n");
                htmlStr.Append("      	  class='font101637'>(Payment Method)</font><font class='font71637'>:<span                                                                                                                              \n");
                htmlStr.Append("      	  style='mso-spacerun:yes'> </span></font></td>                                                                                                                                                         \n");
                htmlStr.Append("      	  <td colspan=10 class=xl2011637>&nbsp;" + dt.Rows[0]["PaymentMethodCK"] + "</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl671637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=22 style='mso-height-source:userset;height:15.0pt'>                                                                                                                                         \n");
                htmlStr.Append("      	  <td height=22 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                                                                       \n");
                htmlStr.Append("      	  <td class=xl701637 colspan=3>Lo&#7841;i ti&#7873;n t&#7879;&nbsp;<font                                                                                                                                \n");
                htmlStr.Append("      	  class='font101637'>(Currency)</font><font class='font71637'>:</font></td>                                                                                                                             \n");
                htmlStr.Append("      	  <td colspan=6 class=xl2021637>&nbsp; " + dt.Rows[0]["CurrencyCodeUSD"] + "</td>                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl651637 colspan=4><font class='font81637'>T&#7927; giá</font><font                                                                                                                         \n");
                htmlStr.Append("      	  class='font91637'>&nbsp;(</font><font class='font111637'>Exchange rate</font><font                                                                                                                    \n");
                htmlStr.Append("      	  class='font91637'>):&nbsp;</font></td>                                                                                                                                                                \n");
                htmlStr.Append("      	  <td colspan=4 class=xl2021637>&nbsp;" + dt.Rows[0]["tr_rate_88"] + "</td>                                                                                                                                             \n");
                htmlStr.Append("      	  <td class=xl671637>&nbsp;</td>                                                                                                                                                                        \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=4 style='mso-height-source:userset;height:3.0pt'>                                                                                                                                           \n");
                htmlStr.Append("      	  <td colspan=19 height=4 class=xl1811637 style='border-right:.5pt solid black;                                                                                                                         \n");
                htmlStr.Append("      	  height:3.0pt'>&nbsp;</td>                                                                                                                                                                             \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=21 style='height:15.0pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td colspan=2 height=21 class=xl1711637 style='height:15.0pt'>STT</td>                                                                                                                                \n");
                htmlStr.Append("      	  <td colspan=7 class=xl1711637 style='border-right:.5pt solid black'>Tên hàng                                                                                                                          \n");
                htmlStr.Append("      	  hóa, d&#7883;ch v&#7909;</td>                                                                                                                                                                         \n");
                htmlStr.Append("      	  <td class=xl1221637 style='border-top:none'>&#272;&#417;n v&#7883; tính</td>                                                                                                                          \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1711637 style='border-right:.5pt solid black'>S&#7889;                                                                                                                          \n");
                htmlStr.Append("      	  l&#432;&#7907;ng</td>                                                                                                                                                                                 \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1221637>&#272;&#417;n giá</td>                                                                                                                                                  \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1711637 style='border-right:.5pt solid black'>Thành                                                                                                                             \n");
                htmlStr.Append("      	  ti&#7873;n</td>                                                                                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=18 style='height:13.2pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td colspan=2 height=18 class=xl1731637 style='height:13.2pt'>No.</td>                                                                                                                                \n");
                htmlStr.Append("      	  <td colspan=7 class=xl1741637 style='border-right:.5pt solid black'>Description</td>                                                                                                                  \n");
                htmlStr.Append("      	  <td class=xl1231637>Unit</td>                                                                                                                                                                         \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1731637 style='border-right:.5pt solid black'>Quantity</td>                                                                                                                     \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1231637>Unit price</td>                                                                                                                                                         \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1731637 style='border-right:.5pt solid black'>Amount</td>                                                                                                                       \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                htmlStr.Append("      	 <tr height=18 style='height:13.8pt'>                                                                                                                                                                   \n");
                htmlStr.Append("      	  <td colspan=2 height=18 class=xl1671637 style='height:13.8pt'>1</td>                                                                                                                                  \n");
                htmlStr.Append("      	  <td colspan=7 class=xl1671637 style='border-right:.5pt solid black'>2</td>                                                                                                                            \n");
                htmlStr.Append("      	  <td class=xl1241637>3</td>                                                                                                                                                                            \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1671637 style='border-right:.5pt solid black'>4</td>                                                                                                                            \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1241637>5</td>                                                                                                                                                                  \n");
                htmlStr.Append("      	  <td colspan=3 class=xl1671637 style='border-right:.5pt solid black'>6 = 4 x 5</td>                                                                                                                    \n");
                htmlStr.Append("      	 </tr>                                                                                                                                                                                                  \n");
                
                
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
                      // htmlStr.Append("                                                                   \n");
                      // htmlStr.Append("    <tr class=xl796926 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "' >                                                               \n");
                      // htmlStr.Append("      <td colspan=2 height=25 class=xl1316926 width=38 style='border-right:.5pt solid black;border-top:.5pt solid black;                                                                 \n");
                      // htmlStr.Append("      height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; width:36.25pt'>" + dt_d.Rows[v_index][7] + "</td>                                                               \n");
                      // htmlStr.Append("      <td colspan=6 class=xl1336926 width=276 style='border-right:.5pt solid black; border-top:.5pt solid black;                                                               \n");
                      // htmlStr.Append("      border-left:none;width:206pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                               \n");
                      // htmlStr.Append("      <td class=xl1176926 style='border-left:none;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][1] + "</td>                                                               \n");
                      // htmlStr.Append("      <td colspan=2 class=xl1806926 style='border-left:none;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][2] + "</td>                                                               \n");
                      // htmlStr.Append("      <td class=xl1146926 style='border-top:.5pt solid black; '>&nbsp;</td>                                                               \n");
                      // htmlStr.Append("      <td colspan=2 class=xl1366926 style='border-left:none;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                      // htmlStr.Append("      <td class=xl1136926 style='border-top:.5pt solid black; '></td>                                                               \n");
                      // htmlStr.Append("      <td colspan=2 class=xl1386926 width=141 style='border-left:none;width:106pt;border-top:.5pt solid black; '>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                      // htmlStr.Append("      <td class=xl1136926 style='border-top:.5pt solid black; '>&nbsp;</td>                                                               \n");
                      // htmlStr.Append("     </tr>                                                               \n");

                        htmlStr.Append("    	 <tr height=25 style='mso-height-source:userset;height:18.05pt'>																										 \n");
                        htmlStr.Append("    	  <td colspan=2 height=25 class=xl1251637 width=39 style='border-right:.5pt solid black;                                                                                 \n");
                        htmlStr.Append("    	  height:18.05pt;width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                                    \n");
                        htmlStr.Append("    	  <td colspan=7 class=xl1531637 width=288 style='border-right:.5pt solid black;                                                                                          \n");
                        htmlStr.Append("    	  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                                                 \n");
                        htmlStr.Append("    	  <td class=xl1031637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                                  \n");
                        htmlStr.Append("    	  <td colspan=2 class=xl1691637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                                                        \n");
                        htmlStr.Append("    	  <td class=xl1001637>&nbsp;</td>                                                                                                                                        \n");
                        htmlStr.Append("    	  <td colspan=2 class=xl1561637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                                                        \n");
                        htmlStr.Append("    	  <td class=xl991637>&nbsp;</td>                                                                                                                                         \n");
                        htmlStr.Append("    	  <td colspan=2 class=xl1581637 width=133 style='border-left:none;width:100pt'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                                  \n");
                        htmlStr.Append("    	  <td class=xl991637>&nbsp;</td>                                                                                                                                         \n");
                        htmlStr.Append("    	 </tr>                                                                                                                                                                   \n");







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
                            htmlStr.Append("      <td colspan=2 class=xl1256926 width=29 style='border-top:none;width:22pt'>" + dt_d.Rows[v_index][2] + "</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1226926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][3] + "</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1236926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("      <td colspan=2 class=xl1956926 style='border-left:none'>" + dt_d.Rows[v_index][4] + "</td>                                                               \n");
                            htmlStr.Append("      <td class=xl1156926 width=6 style='border-top:none;width:6.25pt'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("     </tr>                                                               \n");


                            htmlStr.Append("    	<tr height=25 style='mso-height-source:userset;height:18.05pt'>																			\n");
                            htmlStr.Append("    	  <td colspan=2 height=25 class=xl1491637 width=39 style='border-right:.5pt solid black;                                                \n");
                            htmlStr.Append("    	  height:18.05pt;width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                   \n");
                            htmlStr.Append("    	  <td colspan=7 class=xl1891637 width=288 style='border-right:.5pt solid black;                                                         \n");
                            htmlStr.Append("    	  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                \n");
                            htmlStr.Append("    	  <td class=xl1051637 width=78 style='border-top:none;border-left:none;                                                                 \n");
                            htmlStr.Append("    	  width:58pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                                  \n");
                            htmlStr.Append("    	  <td colspan=2 class=xl1921637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                       \n");
                            htmlStr.Append("    	  <td class=xl1061637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                                             \n");
                            htmlStr.Append("    	  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                       \n");
                            htmlStr.Append("    	  <td class=xl1071637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                                             \n");
                            htmlStr.Append("    	  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                       \n");
                            htmlStr.Append("    	  <td class=xl1011637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                                             \n");
                            htmlStr.Append("    	 </tr>				                                                                                                                    \n");


                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                               /* htmlStr.Append("    <tr class=xl746926 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
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
                               */
                                htmlStr.Append("    	<tr height=25 style='mso-height-source:userset;height:18.05pt'>																			\n");
                                htmlStr.Append("    	  <td colspan=2 height=25 class=xl1491637 width=39 style='border-right:.5pt solid black;                                                \n");
                                htmlStr.Append("    	  height:18.05pt;width:29pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                   \n");
                                htmlStr.Append("    	  <td colspan=7 class=xl1891637 width=288 style='border-right:.5pt solid black;                                                         \n");
                                htmlStr.Append("    	  border-left:none;width:216pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                                                \n");
                                htmlStr.Append("    	  <td class=xl1051637 width=78 style='border-top:none;border-left:none;                                                                 \n");
                                htmlStr.Append("    	  width:58pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                                                                  \n");
                                htmlStr.Append("    	  <td colspan=2 class=xl1921637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                       \n");
                                htmlStr.Append("    	  <td class=xl1061637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                                             \n");
                                htmlStr.Append("    	  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                       \n");
                                htmlStr.Append("    	  <td class=xl1071637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                                             \n");
                                htmlStr.Append("    	  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                       \n");
                                htmlStr.Append("    	  <td class=xl1011637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                                             \n");
                                htmlStr.Append("    	 </tr>				                                                                                                                    \n");

                            }
                            else
                            {
                                /*htmlStr.Append("    <tr class=xl746926 height=25 style='mso-height-source:userset;height:" + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
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
                            */
                                htmlStr.Append("    	 <tr height=25 style='mso-height-source:userset;height:18.05pt'>																		 \n");
                                htmlStr.Append("    	  <td height=25 colspan=2 class=xl1251637 width=6 style='height:18.05pt;border-top:none;border-right:.5pt solid black;                   \n");
                                htmlStr.Append("    	  width:4pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                    \n");
                                htmlStr.Append("    	  <td colspan=7 class=xl1531637 width=288 style='border-right:.5pt solid black;                                                          \n");
                                htmlStr.Append("    	  border-left:none;width:216pt;border-bottom:none'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString()+ "</td>                                                              \n");
                                htmlStr.Append("    	  <td class=xl1041637 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                  \n");
                                htmlStr.Append("    	  <td colspan=2 class=xl1561637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                        \n");
                                htmlStr.Append("    	  <td class=xl1001637 style='border-top:none'>&nbsp;</td>                                                                                \n");
                                htmlStr.Append("    	  <td colspan=2 class=xl1561637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                        \n");
                                htmlStr.Append("    	  <td class=xl991637 style='border-top:none'>&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("    	  <td colspan=2 class=xl1631637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                        \n");
                                htmlStr.Append("    	  <td class=xl991637 style='border-top:none'>&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("    	 </tr>			                                                                                                                         \n");

                            }

                        }
                    }
                    else
                    { // dong giua                                                                                                                                    
                        /*htmlStr.Append("    	<tr class=xl796926 height=25 style='mso-height-source:userset;height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                               \n");
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
                   */
                        htmlStr.Append("    	 <tr height=25 style='mso-height-source:userset;height:18.05pt'>																		 \n");
                        htmlStr.Append("    	  <td height=25 colspan=2 class=xl1251637 width=6 style='height:18.05pt;border-top:none;border-right:.5pt solid black;                   \n");
                        htmlStr.Append("    	  width:4pt'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                                                                    \n");
                        htmlStr.Append("    	  <td colspan=7 class=xl1531637 width=288 style='border-right:.5pt solid black;                                                          \n");
                        htmlStr.Append("    	  border-left:none;width:216pt;border-bottom:none'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                                                              \n");
                        htmlStr.Append("    	  <td class=xl1041637 style='border-top:none;border-left:none'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                  \n");
                        htmlStr.Append("    	  <td colspan=2 class=xl1561637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][2] + "</td>                                                        \n");
                        htmlStr.Append("    	  <td class=xl1001637 style='border-top:none'>&nbsp;</td>                                                                                \n");
                        htmlStr.Append("    	  <td colspan=2 class=xl1561637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][3] + "</td>                                                        \n");
                        htmlStr.Append("    	  <td class=xl991637 style='border-top:none'>&nbsp;</td>                                                                                 \n");
                        htmlStr.Append("    	  <td colspan=2 class=xl1631637 style='border-left:none'>&nbsp;" + dt_d.Rows[v_index][4] + "</td>                                                        \n");
                        htmlStr.Append("    	  <td class=xl991637 style='border-top:none'>&nbsp;</td>                                                                                 \n");
                        htmlStr.Append("    	 </tr>			                                                                                                                         \n");

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
                            /*htmlStr.Append("    <tr class=xl746926 height=25 style='mso-height-source:userset;height:" + v_rowHeightEmptyLast + "'>                                                               \n");
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
                        */
                            htmlStr.Append("    		<tr height=25 style='mso-height-source:userset;height:18.05pt'>															\n");
                            htmlStr.Append("    		  <td colspan=2 height=25 class=xl1491637 width=39 style='border-right:.5pt solid black;                                \n");
                            htmlStr.Append("    		  height:18.05pt;width:29pt'>&nbsp;</td>                                                                                \n");
                            htmlStr.Append("    		  <td colspan=7 class=xl1891637 width=288 style='border-right:.5pt solid black;                                         \n");
                            htmlStr.Append("    		  border-left:none;width:216pt'>&nbsp;</td>                                                                             \n");
                            htmlStr.Append("    		  <td class=xl1051637 width=78 style='border-top:none;border-left:none;                                                 \n");
                            htmlStr.Append("    		  width:58pt'>&nbsp;</td>                                                                                               \n");
                            htmlStr.Append("    		  <td colspan=2 class=xl1921637 style='border-left:none'>&nbsp;</td>                                                    \n");
                            htmlStr.Append("    		  <td class=xl1061637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                             \n");
                            htmlStr.Append("    		  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;</td>                                                    \n");
                            htmlStr.Append("    		  <td class=xl1071637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                             \n");
                            htmlStr.Append("    		  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;</td>                                                    \n");
                            htmlStr.Append("    		  <td class=xl1011637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                             \n");
                            htmlStr.Append("    		 </tr>	                                                                                                                \n");


                        }
                        else
                        {
                            /* htmlStr.Append("    	<tr class=xl796926 height=25 style='mso-height-source:userset;height: " + v_rowHeightEmptyLast + "'>                                                               \n");
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
                      */
                            htmlStr.Append("    			<tr height=25 style='mso-height-source:userset;height:18.05pt'>															\n");
                            htmlStr.Append("    			  <td colspan=2 height=25 class=xl1251637 width=39 style='border-right:.5pt solid black;                                \n");
                            htmlStr.Append("    			  height:18.05pt;width:29pt'>&nbsp;</td>                                                                                \n");
                            htmlStr.Append("    			  <td colspan=7 class=xl1531637 width=288 style='border-right:.5pt solid black;                                         \n");
                            htmlStr.Append("    			  border-left:none;width:216pt'>&nbsp;</td>                                                                             \n");
                            htmlStr.Append("    			  <td class=xl1041637 style='border-top:none;border-left:none'>&nbsp;</td>                                              \n");
                            htmlStr.Append("    			  <td colspan=2 class=xl1561637 style='border-left:none'>&nbsp;</td>                                                    \n");
                            htmlStr.Append("    			  <td class=xl1001637 style='border-top:none'>&nbsp;</td>                                                               \n");
                            htmlStr.Append("    			  <td colspan=2 class=xl1561637 style='border-left:none'>&nbsp;</td>                                                    \n");
                            htmlStr.Append("    			  <td class=xl991637 style='border-top:none'>&nbsp;</td>                                                                \n");
                            htmlStr.Append("    			  <td colspan=2 class=xl1581637 width=133 style='border-left:none;width:100pt'>&nbsp;</td>                              \n");
                            htmlStr.Append("    			  <td class=xl991637 style='border-top:none'>&nbsp;</td>                                                                \n");
                            htmlStr.Append("    			 </tr>                                                                                                                  \n");

                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    // htmlStr.Append("    <tr class=xl746926 height=26 style='mso-height-source:userset;height:" + (v_spacePerPage).ToString() + "pt'>                                                               \n");
                    // htmlStr.Append("      <td height=26 class=xl976926 width=6 style='height:" + (v_spacePerPage).ToString() + "pt;width:6.25pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1246926 width=32 style='width:30pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1246926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1246926 width=53 style='width:40pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1246926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1246926 width=39 style='width:36.25pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1246926 width=67 style='width:62.5pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1246926 width=11 style='width:10pt'>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("      <td colspan=7 class=xl1876926 width=255 style='                                                               \n");
                    // htmlStr.Append("      width:193pt'></td>                                                               \n");
                    // htmlStr.Append("      <td colspan=2 class=xl1896926 style='border-left:none'></td>                                                               \n");
                    // htmlStr.Append("      <td class=xl1166926>&nbsp;</td>                                                               \n");
                    // htmlStr.Append("     </tr>                                                       \n");
                   
                    htmlStr.Append("    		<tr height=25 style='mso-height-source:userset;height:18.05pt'>															\n");
                    htmlStr.Append("    		  <td colspan=2 height=25 class=xl1491637 width=39 style='border-right:.5pt solid black;                                \n");
                    htmlStr.Append("    		  height:18.05pt;width:29pt'>&nbsp;</td>                                                                                \n");
                    htmlStr.Append("    		  <td colspan=7 class=xl1891637 width=288 style='border-right:.5pt solid black;                                         \n");
                    htmlStr.Append("    		  border-left:none;width:216pt'>&nbsp;</td>                                                                             \n");
                    htmlStr.Append("    		  <td class=xl1051637 width=78 style='border-top:none;border-left:none;                                                 \n");
                    htmlStr.Append("    		  width:58pt'>&nbsp;</td>                                                                                               \n");
                    htmlStr.Append("    		  <td colspan=2 class=xl1921637 style='border-left:none'>&nbsp;</td>                                                    \n");
                    htmlStr.Append("    		  <td class=xl1061637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                             \n");
                    htmlStr.Append("    		  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;</td>                                                    \n");
                    htmlStr.Append("    		  <td class=xl1071637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                             \n");
                    htmlStr.Append("    		  <td colspan=2 class=xl1511637 style='border-left:none'>&nbsp;</td>                                                    \n");
                    htmlStr.Append("    		  <td class=xl1011637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                             \n");
                    htmlStr.Append("    		 </tr>	                                                                                                                \n");




                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=18 style='height: 10pt'>                                                                                                                                                                \n");
                    htmlStr.Append("			<td colspan=27 height=18                                                                                                                                                       \n");
                    htmlStr.Append("				style=' height: 10pt'>&nbsp;</td>                                                                                                                           \n");
                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             

            htmlStr.Append("    			<tr height=24 style='mso-height-source:userset;height:18.0pt'>																					 \n");
            htmlStr.Append("    			  <td height=24 class=xl831637 width=6 style='height:18.0pt;width:4pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    			  <td class=xl1081637 width=33 style='width:25pt'>&nbsp;</td>                                                                                    \n");
            htmlStr.Append("    			  <td class=xl1081637 width=70 style='border-top:none;width:52pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1081637 width=55 style='border-top:none;width:41pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1081637 width=34 style='border-top:none;width:26pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1081637 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                                                      \n");
            htmlStr.Append("    			  <td class=xl1081637 width=41 style='border-top:none;width:31pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1081637 width=70 style='border-top:none;width:52pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1081637 width=12 style='border-top:none;width:9pt'>&nbsp;</td>                                                                     \n");
            htmlStr.Append("    			  <td colspan=7 class=xl1431637 width=266 style='border-right:.5pt solid black;                                                                  \n");
            htmlStr.Append("    			  width:199pt'>C&#7897;ng ti&#7873;n hàng (<font class='font111637'>Sub amount</font><font                                                       \n");
            htmlStr.Append("    			  class='font81637'>) :</font></td>                                                                                                              \n");
            htmlStr.Append("    			  <td colspan=2 class=xl1451637 style='border-left:none'>" + amount_trans + "&nbsp;</td>                                                               \n");
            htmlStr.Append("    			  <td class=xl1021637>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                                                  \n");
            htmlStr.Append("    			  <td height=24 class=xl831637 width=6 style='height:18.0pt;border-top:none;                                                                     \n");
            htmlStr.Append("    			  width:4pt'>&nbsp;</td>                                                                                                                         \n");
            htmlStr.Append("    			  <td colspan=4 class=xl1271637 width=192 style='width:144pt'><span                                                                              \n");
            htmlStr.Append("    			  style='mso-spacerun:yes'> </span>Thu&#7871; su&#7845;t GTGT (<font                                                                             \n");
            htmlStr.Append("    			  class='font111637'>Tax rate</font><font class='font81637'>) : </font></td>                                                                     \n");
            htmlStr.Append("    			  <td class=xl1271637 width=6 style='border-top:none;width:5pt'>&nbsp;" + dt.Rows[0]["taxrate"] + "</td>                                                       \n");
            htmlStr.Append("    			  <td class=xl1271637 width=41 style='border-top:none;width:31pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1091637 width=70 style='border-top:none;width:52pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1271637 width=12 style='border-top:none;width:9pt'>&nbsp;</td>                                                                     \n");
            htmlStr.Append("    			  <td colspan=7 class=xl1431637 width=266 style='border-right:.5pt solid black;                                                                  \n");
            htmlStr.Append("    			  width:199pt'>Ti&#7873;n thu&#7871; GTGT (<font class='font111637'>VAT</font><font                                                              \n");
            htmlStr.Append("    			  class='font81637'>) :</font></td>                                                                                                              \n");
            htmlStr.Append("    			  <td colspan=2 class=xl1451637 style='border-left:none'>" + amount_vat + "&nbsp;</td>                                                             \n");
            htmlStr.Append("    			  <td class=xl731637 style='border-top:none'>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=24 style='mso-height-source:userset;height:18.0pt'>                                                                                  \n");
            htmlStr.Append("    			  <td height=24 class=xl831637 width=6 style='height:18.0pt;border-top:none;                                                                     \n");
            htmlStr.Append("    			  width:4pt'>&nbsp;</td>                                                                                                                         \n");
            htmlStr.Append("    			  <td class=xl1271637 width=33 style='border-top:none;width:25pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1271637 width=70 style='border-top:none;width:52pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1271637 width=55 style='border-top:none;width:41pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1271637 width=34 style='border-top:none;width:26pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1271637 width=6 style='border-top:none;width:5pt'>&nbsp;</td>                                                                      \n");
            htmlStr.Append("    			  <td class=xl1271637 width=41 style='border-top:none;width:31pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1271637 width=70 style='border-top:none;width:52pt'>&nbsp;</td>                                                                    \n");
            htmlStr.Append("    			  <td class=xl1271637 width=12 style='border-top:none;width:9pt'>&nbsp;</td>                                                                     \n");
            htmlStr.Append("    			  <td colspan=7 class=xl1431637 width=266 style='border-right:.5pt solid black;                                                                  \n");
            htmlStr.Append("    			  width:199pt'>T&#7893;ng ti&#7873;n thanh toán (<font class='font111637'>Total                                                                  \n");
            htmlStr.Append("    			  payment</font><font class='font81637'>) :</font></td>                                                                                          \n");
            htmlStr.Append("    			  <td colspan=2 class=xl1451637 style='border-left:none'>" + amount_total + "&nbsp;</td>                                                      \n");
            htmlStr.Append("    			  <td class=xl731637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=21 style='height:15.0pt'>                                                                                                            \n");
            htmlStr.Append("    			  <td height=21 class=xl721637 style='height:15.0pt;border-top:none'>&nbsp;</td>                                                                 \n");
            htmlStr.Append("    			  <td colspan=17 class=xl1471637>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng                                                                       \n");
            htmlStr.Append("    			  ch&#7919; (<font class='font111637'>In words</font><font class='font71637'>):                                                                  \n");
            htmlStr.Append("    			   <%=v_read_vie%></font></td>                                                                                                                   \n");
            htmlStr.Append("    			  <td class=xl741637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=13 style='mso-height-source:userset;height:10.05pt'>                                                                                 \n");
            htmlStr.Append("    			  <td colspan=19 height=13 class=xl1811637 style='border-right:.5pt solid black;                                                                 \n");
            htmlStr.Append("    			  height:10.05pt'>&nbsp;</td>                                                                                                                    \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                  \n");
            htmlStr.Append("    			  <td height=21 class=xl761637 style='height:15.0pt'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("    			  <td colspan=6 class=xl1481637 width=239 style='width:180pt'>Ng&#432;&#7901;i                                                                   \n");
            htmlStr.Append("    			  mua hàng (Buyer)</td>                                                                                                                          \n");
            htmlStr.Append("    			  <td colspan=4 class=xl1481637 width=216 style='width:161pt'>&nbsp;</td>                                                                        \n");
            htmlStr.Append("    			  <td colspan=7 class=xl1481637 width=265 style='width:199pt'>Ng&#432;&#7901;i                                                                   \n");
            htmlStr.Append("    			  mua hàng (<font class='font141637'>Seller</font><font class='font51637'>)</font></td>                                                          \n");
            htmlStr.Append("    			  <td class=xl751637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                  \n");
            htmlStr.Append("    			  <td height=21 class=xl1171637 style='height:15.0pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("    			  <td colspan=6 class=xl1401637 width=239 style='width:180pt'>(Ký, ghi rõ                                                                        \n");
            htmlStr.Append("    			  h&#7885; tên)</td>                                                                                                                             \n");
            htmlStr.Append("    			  <td colspan=4 class=xl1411637 width=216 style='width:161pt'>&nbsp;</td>                                                                        \n");
            htmlStr.Append("    			  <td colspan=7 class=xl1411637 width=265 style='width:199pt'>(Ký, &#273;óng                                                                     \n");
            htmlStr.Append("    			  d&#7845;u, ghi rõ h&#7885; tên)</td>                                                                                                           \n");
            htmlStr.Append("    			  <td class=xl771637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=21 style='mso-height-source:userset;height:15.0pt'>                                                                                  \n");
            htmlStr.Append("    			  <td height=21 class=xl781637 style='height:15.0pt'>&nbsp;</td>                                                                                 \n");
            htmlStr.Append("    			  <td colspan=6 class=xl1421637 width=239 style='width:180pt'>(Signature &amp;                                                                   \n");
            htmlStr.Append("    			  full name)</td>                                                                                                                                \n");
            htmlStr.Append("    			  <td colspan=4 class=xl1421637 width=216 style='width:161pt'>&nbsp;</td>                                                                        \n");
            htmlStr.Append("    			  <td colspan=7 class=xl1421637 width=265 style='width:199pt'>(Signature, stamp                                                                  \n");
            htmlStr.Append("    			  &amp; full name)</td>                                                                                                                          \n");
            htmlStr.Append("    			  <td class=xl791637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=18 style='mso-height-source:userset;height:13.95pt'>                                                                                 \n");
            htmlStr.Append("    			  <td height=18 class=xl1191637 style='height:13.95pt'>&nbsp;</td>                                                                               \n");
            htmlStr.Append("    			  <td class=xl681637 width=33 style='width:25pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=70 style='width:52pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=55 style='width:41pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=34 style='width:26pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=6 style='width:5pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			  <td class=xl681637 width=41 style='width:31pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=70 style='width:52pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("    			  <td class=xl681637 width=78 style='width:58pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=56 style='width:42pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=30 style='width:23pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			  <td class=xl681637 width=41 style='width:31pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=49 style='width:37pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			  <td class=xl681637 width=41 style='width:31pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=92 style='width:69pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl801637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=22 style='height:16.8pt'>                                                                                                            \n");
            htmlStr.Append("    			  <td height=22 class=xl1191637 style='height:16.8pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl681637 width=34 style='width:26pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=6 style='width:5pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			  <td class=xl681637 width=41 style='width:31pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=70 style='width:52pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("    			  <td class=xl681637 width=78 style='width:58pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("    			  <td class=xl1121637 colspan=3>Signature Valid</td>                                                                                             \n");
                htmlStr.Append("    			  <td align=left valign=top class=xl861637><![if !vml]><span style='mso-ignore:vglayout;                                                         \n");
                htmlStr.Append("    			  position:absolute;z-index:1;margin-left:22px;margin-top:18px;width:33px;                                                                       \n");
                htmlStr.Append("    			  height:50px'><img width=33 height=50                                                                                                           \n");
                htmlStr.Append("    			  src='${pageContext.request.contextPath}/assets/img/check_signed.png'                                                                           \n");
                htmlStr.Append("    			  v:shapes='Picture_x0020_8'></span><![endif]><span style='mso-ignore:vglayout2'>                                                                \n");
                htmlStr.Append("    			  <table cellpadding=0 cellspacing=0>                                                                                                            \n");
                htmlStr.Append("    			   <tr>                                                                                                                                          \n");
                htmlStr.Append("    				<td height=22  width=49 style='height:16.8pt;width:37pt'>&nbsp;</td>                                                                         \n");
                htmlStr.Append("    			   </tr>                                                                                                                                         \n");
                htmlStr.Append("    			  </table>                                                                                                                                       \n");
                htmlStr.Append("    			  </span></td>                                                                                                                                   \n");
                htmlStr.Append("    			  <td class=xl871637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl871637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl881637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl901637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			 </tr>                                                                                                                                           \n");
                htmlStr.Append("    			 <tr height=40 style='mso-height-source:userset;height:30.0pt'>                                                                                  \n");
                htmlStr.Append("    			  <td height=40 class=xl931637 style='height:30.0pt'>&nbsp;</td>                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td colspan=7 class=xl1351637 width=265 style='border-right:.5pt solid black;                                                                  \n");
                htmlStr.Append("    			  width:199pt'>&#272;&#432;&#7907;c ký b&#7903;i: <font class='font161637'><%=l_company_nm %></font></td>                                        \n");

            }
            else
            {

                htmlStr.Append("    				  <td class=xl1121637 colspan=3></td>                                                                                                        \n");
                htmlStr.Append("    			  <td align=left valign=top class=xl861637></td>                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl871637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl871637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl881637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl901637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			 </tr>                                                                                                                                           \n");
                htmlStr.Append("    			 <tr height=40 style='mso-height-source:userset;height:30.0pt'>                                                                                  \n");
                htmlStr.Append("    			  <td height=40 class=xl931637 style='height:30.0pt'>&nbsp;</td>                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl951637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td class=xl941637>&nbsp;</td>                                                                                                                 \n");
                htmlStr.Append("    			  <td colspan=7 class=xl1351637 width=265 style='border-right:.5pt solid black;                                                                  \n");
                htmlStr.Append("    			  width:199pt'>&#272;&#432;&#7907;c ký b&#7903;i:</td>                                                                                           \n");

            }

            htmlStr.Append("    			  <td class=xl961637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=21 style='height:15.0pt'>                                                                                                            \n");
            htmlStr.Append("    			  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("    			  <td class=xl661637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl681637 width=34 style='width:26pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=6 style='width:5pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			  <td class=xl681637 width=41 style='width:31pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=70 style='width:52pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=12 style='width:9pt'>&nbsp;</td>                                                                                      \n");
            htmlStr.Append("    			  <td class=xl681637 width=78 style='width:58pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl1111637 colspan=3>Ngày ký:<span style='mso-spacerun:yes'> </span>" + dt.Rows[0]["SignedBy"] + "</td>                                               \n");
            htmlStr.Append("    			  <td class=xl1151637>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("    			  <td class=xl1151637>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("    			  <td class=xl1151637>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("    			  <td class=xl1161637>&nbsp;</td>                                                                                                                \n");
            htmlStr.Append("    			  <td class=xl891637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			                                                                                                                                                 \n");
            htmlStr.Append("    			 <tr height=21 style='height:15.0pt'>                                                                                                            \n");
            htmlStr.Append("    			  <td height=21 class=xl1191637 style='height:15.0pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("    			  <td class=xl651637 colspan=9>Tra c&#7913;u t&#7841;i Website: <font                                                                            \n");
            htmlStr.Append("    			  class='font61637'><span style='mso-spacerun:yes'> </span></font><font                                                                          \n");
            htmlStr.Append("    			  class='font121637'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></td>                                                                            \n");
            htmlStr.Append("    			  <td class=xl651637>&nbsp;</td>                                                                                                                 \n");
            htmlStr.Append("    			  <td class=xl921637 colspan=4>Mã nh&#7853;n hóa &#273;&#417;n:  " + dt.Rows[0]["SignedDate"] + "</td>                                                           \n");
            htmlStr.Append("    			  <td class=xl681637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			  <td class=xl681637 width=41 style='width:31pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl681637 width=92 style='width:69pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("    			  <td class=xl801637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=21 style='height:15.0pt'>                                                                                                            \n");
            htmlStr.Append("    			  <td height=21 class=xl1281637 style='height:15.0pt'>&nbsp;</td>                                                                                \n");
            htmlStr.Append("    			  <td colspan=17 class=xl1381637>(C&#7847;n ki&#7875;m tra, &#273;&#7889;i                                                                       \n");
            htmlStr.Append("    			  chi&#7871;u khi l&#7853;p, giao nh&#7853;n hóa &#273;&#417;n)</td>                                                                             \n");
            htmlStr.Append("    			  <td class=xl821637 width=6 style='width:4pt'>&nbsp;</td>                                                                                       \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <tr height=21 style='height:15.0pt'>                                                                                                            \n");
            htmlStr.Append("    			  <td height=21 class=xl1201637 style='height:15.0pt;border-top:none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("    			  <td colspan=17 class=xl1391637>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                                               \n");
            htmlStr.Append("    			  <td class=xl1101637 width=6 style='border-top:none;width:4pt'>&nbsp;</td>                                                                      \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <![if supportMisalignedColumns]>                                                                                                                \n");
            htmlStr.Append("    			 <tr height=0 style='display:none'>                                                                                                              \n");
            htmlStr.Append("    			  <td width=6 style='width:4pt'></td>                                                                                                            \n");
            htmlStr.Append("    			  <td width=33 style='width:25pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=70 style='width:52pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=55 style='width:41pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=34 style='width:26pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=6 style='width:5pt'></td>                                                                                                            \n");
            htmlStr.Append("    			  <td width=41 style='width:31pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=70 style='width:52pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=12 style='width:9pt'></td>                                                                                                           \n");
            htmlStr.Append("    			  <td width=78 style='width:58pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=56 style='width:42pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=30 style='width:23pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=6 style='width:4pt'></td>                                                                                                            \n");
            htmlStr.Append("    			  <td width=41 style='width:31pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=49 style='width:37pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=6 style='width:4pt'></td>                                                                                                            \n");
            htmlStr.Append("    			  <td width=41 style='width:31pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=92 style='width:69pt'></td>                                                                                                          \n");
            htmlStr.Append("    			  <td width=6 style='width:4pt'></td>                                                                                                            \n");
            htmlStr.Append("    			 </tr>                                                                                                                                           \n");
            htmlStr.Append("    			 <![endif]>                                                                                                                                      \n");
            htmlStr.Append("    			</table>					                                                                                                                     \n"); htmlStr.Append("   </ body >          \n");
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