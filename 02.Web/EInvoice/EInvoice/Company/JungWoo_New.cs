using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
//using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;
//using System.Data.OracleClient;
using System.IO;
using System.Text;

namespace EInvoice.Company
{
    public class JungWoo_New
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
            string read_prive = "", read_en = "", read_amount = "", amount_vat = "", amount_total = "", amount_trans = "", amount_net = "", lb_amount_trans = "";


            if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            {
                lb_amount_trans = "";
                amount_trans = "";
                amount_total = dt.Rows[0]["TOTALAMOUNT_DISPLAY_VIE"].ToString();
                amount_vat = dt.Rows[0]["VAT_TR_AMT_USD_DISPLAY"].ToString();
                amount_net = dt.Rows[0]["TOTAL_TR_AMT_VIE"].ToString();
                // read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmountInWord"].ToString()));
            }
            else
            {
                lb_amount_trans = "Tổng cộng VND <font class='font1127974'>(Amount VND):</font>";
                amount_trans = dt.Rows[0]["AMT_BK_VIE_DISPLAY"].ToString();
                amount_total = dt.Rows[0]["netamount_display"].ToString();
                amount_vat = dt.Rows[0]["vat_tr_amt_vie_display"].ToString();
                amount_net = dt.Rows[0]["TotalAmountInWord"].ToString();

                // read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            }

            if (dt.Rows[0]["vatamount_display"].ToString().Trim() == "0")
            {
                amount_vat = "-";
            }
            else
            {
                amount_vat = dt.Rows[0]["vatamount_display"].ToString();
            }
            read_prive = dt.Rows[0]["AMOUNT_WORD_VIE"].ToString();//
            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";

            string l_tax_code = dt.Rows[0]["SELLER_TAXCODE"].ToString();

            string l_taxcode_0 = l_tax_code.Substring(0, 1);
            string l_taxcode_1 = l_tax_code.Substring(1, 1);
            string l_taxcode_2 = l_tax_code.Substring(2, 1);
            string l_taxcode_3 = l_tax_code.Substring(3, 1);
            string l_taxcode_4 = l_tax_code.Substring(4, 1);
            string l_taxcode_5 = l_tax_code.Substring(5, 1);
            string l_taxcode_6 = l_tax_code.Substring(6, 1);
            string l_taxcode_7 = l_tax_code.Substring(7, 1);
            string l_taxcode_8 = l_tax_code.Substring(8, 1);
            string l_taxcode_9 = l_tax_code.Substring(9, 1);

            string l_buyer_taxcode = dt.Rows[0]["BUYERTAXCODE"].ToString();
            string l_buyertax_0 = "";
            string l_buyertax_1 = "";
            string l_buyertax_2 = "";
            string l_buyertax_3 = "";
            string l_buyertax_4 = "";
            string l_buyertax_5 = "";
            string l_buyertax_6 = "";
            string l_buyertax_7 = "";
            string l_buyertax_8 = "";
            string l_buyertax_9 = "";
            string l_buyertax_10 = "";
            string l_buyertax_11 = "";
            string l_buyertax_12 = "";
            string l_buyertax_13 = "";

            if (l_buyer_taxcode.Length == 10)
            {
                l_buyertax_0 = l_buyer_taxcode.Substring(0, 1);
                l_buyertax_1 = l_buyer_taxcode.Substring(1, 1);
                l_buyertax_2 = l_buyer_taxcode.Substring(2, 1);
                l_buyertax_3 = l_buyer_taxcode.Substring(3, 1);
                l_buyertax_4 = l_buyer_taxcode.Substring(4, 1);
                l_buyertax_5 = l_buyer_taxcode.Substring(5, 1);
                l_buyertax_6 = l_buyer_taxcode.Substring(6, 1);
                l_buyertax_7 = l_buyer_taxcode.Substring(7, 1);
                l_buyertax_8 = l_buyer_taxcode.Substring(8, 1);
                l_buyertax_9 = l_buyer_taxcode.Substring(9, 1);

            }
            else if (l_buyer_taxcode.Length > 10)
            {
                l_buyertax_0 = l_buyer_taxcode.Substring(0, 1);
                l_buyertax_1 = l_buyer_taxcode.Substring(1, 1);
                l_buyertax_2 = l_buyer_taxcode.Substring(2, 1);
                l_buyertax_3 = l_buyer_taxcode.Substring(3, 1);
                l_buyertax_4 = l_buyer_taxcode.Substring(4, 1);
                l_buyertax_5 = l_buyer_taxcode.Substring(5, 1);
                l_buyertax_6 = l_buyer_taxcode.Substring(6, 1);
                l_buyertax_7 = l_buyer_taxcode.Substring(7, 1);
                l_buyertax_8 = l_buyer_taxcode.Substring(8, 1);
                l_buyertax_9 = l_buyer_taxcode.Substring(9, 1);
                l_buyertax_10 = l_buyer_taxcode.Substring(10, 1);
                if (l_buyer_taxcode.Length < 13)
                {
                    l_buyertax_11 = l_buyer_taxcode.Substring(11, 1);
                }
                else if (l_buyer_taxcode.Length < 14)
                {
                    l_buyertax_11 = l_buyer_taxcode.Substring(11, 1);
                    l_buyertax_12 = l_buyer_taxcode.Substring(12, 1);
                }
                else if (l_buyer_taxcode.Length < 15)
                {
                    l_buyertax_11 = l_buyer_taxcode.Substring(11, 1);
                    l_buyertax_12 = l_buyer_taxcode.Substring(12, 1);
                    l_buyertax_13 = l_buyer_taxcode.Substring(13, 1);
                }

            }


            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																		\n");
            htmlStr.Append("<html>                                                                                                                                                                      \n");
            htmlStr.Append("<head>                                                                                                                                                                      \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                         \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("<script type='text/javascript'                                                                                                                                              \n");
            htmlStr.Append("	src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                                                                 \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                             \n");
            htmlStr.Append("<!-- Normalize or reset CSS with your favorite library -->                                                                                                                  \n");
            //htmlStr.Append("<link rel='stylesheet'                                                                                                                                                      \n");
            //htmlStr.Append("	href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'>                                                                                            \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("<!-- Load paper.css for happy printing -->                                                                                                                                  \n");
            //htmlStr.Append("<link rel='stylesheet'                                                                                                                                                      \n");
            //htmlStr.Append("	href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>                                                                                                \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("<!-- Set page size here: A5, A4 or A3 -->                                                                                                                                   \n");
            htmlStr.Append("<!-- Set also 'landscape' if you need -->                                                                                                                                   \n");
            htmlStr.Append("<style>                                                                                                                                                                     \n");
            htmlStr.Append("@page {                                                                                                                                                                     \n");
            htmlStr.Append("	size: A4                                                                                                                                                                \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("</style>                                                                                                                                                                    \n");
            htmlStr.Append("<link href='https://fonts.googleapis.com/css?family=Tangerine:700'                                                                                                          \n");
            htmlStr.Append("	rel='stylesheet' type='text/css'>                                                                                                                                       \n");
            htmlStr.Append("<style>                                                                                                                                                                     \n");
            htmlStr.Append("/*body   { font-family: serif }                                                                                                                                             \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                         \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                         \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                            \n");
            htmlStr.Append("    h2 + p { font-size: 16.2pt; line-height: 7mm }                                                                                                                          \n");
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
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("<!--                                                                                                                                                                        \n");
            htmlStr.Append("table {                                                                                                                                                                     \n");
            htmlStr.Append("	mso-displayed-decimal-separator: '\\.';                                                                                                                                  \n");
            htmlStr.Append("	mso-displayed-thousand-separator: '\\,';                                                                                                                                 \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font513844 {                                                                                                                                                               \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font613844 {                                                                                                                                                               \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font713844 {                                                                                                                                                               \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 11pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font813844 {                                                                                                                                                               \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font913844 {                                                                                                                                                               \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1013844 {                                                                                                                                                              \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1113844 {                                                                                                                                                              \n");
            htmlStr.Append("	color: #0066CC;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".font1113845 {                                                                                                                                                              \n");
            htmlStr.Append("	color: black;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl1513844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Calibri;                                                                                                                                                   \n");
            htmlStr.Append("	mso-generic-font-family: auto;                                                                                                                                          \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl15138441 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Calibri;                                                                                                                                                   \n");
            htmlStr.Append("	mso-generic-font-family: auto;                                                                                                                                          \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6313844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6413844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6513844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6613844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6713844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 18pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6813844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl6913844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7013844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7113844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7213844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7313844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: Standard;                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7413844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0\\\\.00_\\)\\;\\\\\\(\\#\\\\,\\#\\#0\\.00\\\\\\)';                                                                                                           \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7513844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 2.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7613844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 12.6pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7713844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7813844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: gray;                                                                                                                                                            \n");
            htmlStr.Append("	font-size: 2.0 pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl7913844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");

            htmlStr.Append(".xl7913845 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8013844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: #CCFFCC;                                                                                                                                                    \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8113844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8213844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8313844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8413844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8513844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8613844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 9.9pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8713844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: #0070C0;                                                                                                                                                         \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8813844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: windowtext;                                                                                                                                                      \n");
            htmlStr.Append("	font-size: 11.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl8913844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9013844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9113844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: bottom;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9213844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.9pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                          \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9313844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.9pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                          \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9413844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 9.9pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("	mso-text-control: shrinktofit;                                                                                                                                          \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9513844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9613844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9713844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9813844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl9913844 {                                                                                                                                                                \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0';                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid gray;                                                                                                                                            \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10013844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0';                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid gray;                                                                                                                                            \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10113844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0';                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid gray;                                                                                                                                            \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10213844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10313844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: nowrap;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10413844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10513844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10613844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10713844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0';                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid gray;                                                                                                                                         \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10813844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0';                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid gray;                                                                                                                                         \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl10913844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0';                                                                                                                                         \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: 1pt solid gray;                                                                                                                                         \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11013844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11113844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11213844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: 0%;                                                                                                                                                  \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11313844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11413844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\#\\\\,\\#\\#0\\\\.00_\\)\\;\\\\\\(\\#\\\\,\\#\\#0\\\\.00\\\\\\)';                                                                                                           \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11513844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11613844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 8.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: Standard;                                                                                                                                            \n");
            htmlStr.Append("	text-align: right;                                                                                                                                                      \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11713844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 1.0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 7.5pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11713845 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 1.0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append(".xl11813844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl11913844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12013844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12113844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: gray;                                                                                                                                                            \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: none;                                                                                                                                                    \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12213844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: gray;                                                                                                                                                            \n");
            htmlStr.Append("	font-size: 2.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: none;                                                                                                                                                       \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12313844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.75pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: 'Times New Roman', serif;                                                                                                                                  \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12413844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: italic;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: General;                                                                                                                                             \n");
            htmlStr.Append("	text-align: center;                                                                                                                                                     \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12513844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 10.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 700;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: general;                                                                                                                                                    \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	mso-background-source: auto;                                                                                                                                            \n");
            htmlStr.Append("	mso-pattern: auto;                                                                                                                                                      \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12613844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: red;                                                                                                                                                             \n");
            htmlStr.Append("	font-size: 16.0pt;                                                                                                                                                      \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border: 1pt solid windowtext;                                                                                                                                          \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12713844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: 1pt solid windowtext;                                                                                                                                     \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12813844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: none;                                                                                                                                                     \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append(".xl12913844 {                                                                                                                                                               \n");
            htmlStr.Append("	padding: 0px;                                                                                                                                                           \n");
            htmlStr.Append("	mso-ignore: padding;                                                                                                                                                    \n");
            htmlStr.Append("	color: black;                                                                                                                                                           \n");
            htmlStr.Append("	font-size: 9.0pt;                                                                                                                                                       \n");
            htmlStr.Append("	font-weight: 400;                                                                                                                                                       \n");
            htmlStr.Append("	font-style: normal;                                                                                                                                                     \n");
            htmlStr.Append("	text-decoration: none;                                                                                                                                                  \n");
            htmlStr.Append("	font-family: Arial, sans-serif;                                                                                                                                         \n");
            htmlStr.Append("	mso-font-charset: 0;                                                                                                                                                    \n");
            htmlStr.Append("	mso-number-format: '\\@';                                                                                                                                                \n");
            htmlStr.Append("	text-align: left;                                                                                                                                                       \n");
            htmlStr.Append("	vertical-align: middle;                                                                                                                                                 \n");
            htmlStr.Append("	border-top: 1pt solid windowtext;                                                                                                                                      \n");
            htmlStr.Append("	border-right: 1pt solid windowtext;                                                                                                                                    \n");
            htmlStr.Append("	border-bottom: 1pt solid windowtext;                                                                                                                                   \n");
            htmlStr.Append("	border-left: none;                                                                                                                                                      \n");
            htmlStr.Append("	background: white;                                                                                                                                                      \n");
            htmlStr.Append("	mso-pattern: black none;                                                                                                                                                \n");
            htmlStr.Append("	white-space: normal;                                                                                                                                                    \n");
            htmlStr.Append("}                                                                                                                                                                           \n");
            htmlStr.Append("-->                                                                                                                                                                         \n");
            htmlStr.Append("</style>                                                                                                                                                                    \n");
            htmlStr.Append("                                                                                                                                                                            \n");
            htmlStr.Append("</head>                                                                                                                                                                     \n");
            htmlStr.Append("<body class='A4'>                                                                                                                                                           \n");

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

            double v_totalHeightLastPage = 243.5;// 258.5;

            double v_totalHeightPage = 580;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 540;// 540;

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

                htmlStr.Append("	<div id='Red invoice - domestic_27974' align=center                                                                                                                     \n");
                htmlStr.Append("		x:publishsource='Excel'>                                                                                                                                            \n");
                htmlStr.Append("		<table border=0 cellpadding=0 cellspacing=0 width=715                                                                                                               \n");
                htmlStr.Append("			style='border-collapse: collapse; table-layout: fixed; width: 647pt'>                                                                                        \n");
                htmlStr.Append("			<col width=30                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 1080; width: 25.875pt'> 		<!--1  -->                                                                      \n");
                htmlStr.Append("			<col width=99                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 3527; width: 87pt'> 		<!--2  66.6pt  6-->                                                             \n");
                htmlStr.Append("			<col width=16                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 568; width: 20.5pt'> 		<!--3  10.8pt-->                                                                \n");
                htmlStr.Append("			<col width=90                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 3185; width: 81.625pt'> 		<!--4  60.3pt  -->                                                              \n");
                htmlStr.Append("			<col width=70                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 2503; width: 68.625pt'> 		<!--5  -->                                                                      \n");
                htmlStr.Append("			<col width=25                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 967; width: 21.25pt'> 		<!--6  -->                                                                      \n");
                htmlStr.Append("			<col width=25 span=5                                                                                                                                            \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 824; width: 21.125pt'> 		<!--7,8,9,10,11  15.3pt -->                                                     \n");
                htmlStr.Append("			<col width=13                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 455; width: 24pt'> 			<!--12  9pt-->                                                              \n");
                htmlStr.Append("			<col width=10                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 369; width: 12.75pt'> 		<!--13  7.2pt-->                                                                \n");
                htmlStr.Append("			<col width=23 span=2                                                                                                                                            \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 824; width: 19.125pt'> 		<!--14,15   -->                                                                 \n");
                htmlStr.Append("			<col width=13 span=2                                                                                                                                            \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 455; width: 11.25pt'> 			<!--16,17  -->                                                                  \n");
                htmlStr.Append("			<col width=23                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 824; width: 19.125pt'> 		<!--18    -->                                                                   \n");
                htmlStr.Append("			<col width=13 span=3                                                                                                                                            \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 455; width: 11.25pt'> 			<!--19,20,21  -->                                                               \n");
                htmlStr.Append("			<col width=10                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 369; width: 9pt'> 		<!--22,23,24  -->                                                               \n");
                htmlStr.Append("			<col width=23 span=3                                                                                                                                            \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 824; width: 21.625pt'> 		<!--25  15.3-->                                                                 \n");
                htmlStr.Append("			<col width=26                                                                                                                                                   \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 910; width: 25.125pt'> 		<!--26  -->                                                                     \n");
                htmlStr.Append("			<col width=6                                                                                                                                                    \n");
                htmlStr.Append("				style='mso-width-source: userset; mso-width-alt: 199; width: 4.5pt'> 		<!--27  -->                                                                     \n");
                htmlStr.Append("			<tr height=6 style='mso-height-source: userset; height: 5.57pt'>                                                                                               \n");
                htmlStr.Append("				<td height=6 class=xl1513844 width=30                                                                                                                       \n");
                htmlStr.Append("					style='height: 5.57pt; width: 20.7pt'></td>                                                                                                            \n");
                htmlStr.Append("				<td class=xl1513844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844 width=16 style='width: 10.8pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td colspan=15 rowspan=4 class=xl6713844 width=420                                                                                                          \n");
                htmlStr.Append("					style='width: 393pt;font-size: 16.0pt;'>HÓA &#272;&#416;N<br> <span                                                                                                        \n");
                htmlStr.Append("					style='mso-spacerun: yes'> </span>GIÁ TR&#7882; GIA T&#258;NG                                                                                           \n");
                htmlStr.Append("				</td>                                                                                                                                                       \n");
                htmlStr.Append("				<td class=xl7613844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6713844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6713844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6713844 width=10 style='width: 7.2pt'></td>                                                                                                     \n");
                htmlStr.Append("				<td class=xl6713844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844 width=26 style='width: 17.1pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844 width=6 style='width: 3.6pt'></td>                                                                                                      \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr style='mso-height-source: userset; height: 17.25pt;'>                                                                                               \n");
                htmlStr.Append("				<td height=5 class=xl1513844 style='height: 3.375pt'></td>                                                                                                  \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td colspan=8 rowspan=2 class=xl8113844 width=144                                                                                                           \n");
                htmlStr.Append("					style='width: 108pt'><span                                                                                                           \n");
                htmlStr.Append("					style='mso-spacerun: yes'></span><font class='font513844'></font></td>                                                               \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=14 style='mso-height-source: userset; height: 9.45pt'>                                                                                               \n");
                htmlStr.Append("				<td colspan=2 rowspan=5 height=69 width=129                                                                                                                 \n");
                htmlStr.Append("					style='height: 53.25pt; width: 97pt' align=left valign=top><![if !vml]><span                                                                            \n");
                htmlStr.Append("					style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: 4px; margin-top: 0px; width: 177.5px; height: 86.25px'><img                        \n");
                htmlStr.Append("						width=177.5 height=86.25                                                                                                                                 \n");
                htmlStr.Append("						src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\JUNGWOO_001.png'                                                                              \n");
                htmlStr.Append("						v:shapes='Picture_x0020_1'></span>                                                                                                                  \n");
                htmlStr.Append("				<![endif]><span style='mso-ignore: vglayout2'>                                                                                                              \n");
                htmlStr.Append("						<table cellpadding=0 cellspacing=0>                                                                                                                 \n");
                htmlStr.Append("							<tr>                                                                                                                                            \n");
                htmlStr.Append("								<td colspan=2 rowspan=5 height=69 class=xl12313844 width=129                                                                                \n");
                htmlStr.Append("									style='height: 53.25pt; width: 97pt'></td>                                                                                              \n");
                htmlStr.Append("							</tr>                                                                                                                                           \n");
                htmlStr.Append("						</table>                                                                                                                                            \n");
                htmlStr.Append("				</span></td>                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr style='mso-height-source: userset; height: 17.25pt;'>                                                                                             \n");
                htmlStr.Append("				<td class=xl1513844 style='height: 18.00pt;'></td>                                                                                                \n");
                htmlStr.Append("				<td colspan=8 class=xl8113844 width=144 style='width: 135pt'>Ký                                                                                             \n");
                htmlStr.Append("					hi&#7879;u:<span style='mso-spacerun: yes'>    </span><font                                                                                             \n");
                htmlStr.Append("					class='font513844'>" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["INVOICESERIALNO"] + "</font>                                                                                                             \n");
                htmlStr.Append("				</td>                                                                                                                                                       \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=2 style='mso-height-source: userset; height: 1.5pt'>                                                                                                 \n");
                htmlStr.Append("				<td height=2 class=xl1513844 style='height: 1.5pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl7613844 width=10 style='width: 7.2pt'></td>                                                                                                     \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl7613844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl7613844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl7613844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl7613844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl7613844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=18 style='mso-height-source: userset; height: 17.25pt;'>                                                                                             \n");
                htmlStr.Append("				<td height=18 class=xl1513844 style='height: 17.25pt;'></td>                                                                                                \n");
                htmlStr.Append("				<td colspan=15 class=xl12413844 width=420 style='width: 314pt'></td>                                                                                        \n");
                htmlStr.Append("				<td colspan=8 class=xl12513844 width=144 style='width: 135pt;color: black; font-weight: 400'>S&#7889;:<span style='mso-spacerun: yes; color:red;font-weight:700'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + dt.Rows[0]["InvoiceNumber"] + "</span></td> \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=20.25 style='mso-height-source: userset; height: 20.25pt;'>                                                                                             \n");
                htmlStr.Append("				<td height=17 class=xl1513844 style='height: 20.25pt;'></td>                                                                                                \n");
                htmlStr.Append("				<td colspan=15 rowspan=2 class=xl6813844 width=420                                                                                                          \n");
                htmlStr.Append("					style='width: 392pt'>Ngày<span style='mso-spacerun: yes'>&nbsp;&nbsp;" + dt.Rows[0]["INVOICEISSUEDDATE_DD"] + "                                                                                 \n");
                htmlStr.Append("				</span>tháng<span style='mso-spacerun: yes'> " + dt.Rows[0]["INVOICEISSUEDDATE_MM"] + " </span>n&#259;m<span                                                                            \n");
                htmlStr.Append("					style='mso-spacerun: yes'> " + dt.Rows[0]["INVOICEISSUEDDATE_YYYY"] + " </span></td>                                                                                                \n");
                htmlStr.Append("				<td class=xl7713844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6813844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6813844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6813844 width=10 style='width: 7.2pt'></td>                                                                                                     \n");
                htmlStr.Append("				<td class=xl6813844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=2 style='mso-height-source: userset; height: 2.25pt'>                                                                                                \n");
                htmlStr.Append("				<td height=2 class=xl1513844 style='height: 2.25pt'></td>                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl7713844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6813844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6813844 width=13 style='width: 9pt'></td>                                                                                                       \n");
                htmlStr.Append("				<td class=xl6813844 width=10 style='width: 7.2pt'></td>                                                                                                     \n");
                htmlStr.Append("				<td class=xl6813844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=12 style='mso-height-source: userset; height: 4.0pt'>                                                                                                \n");
                htmlStr.Append("				<td colspan=26 height=12 class=xl12213844 width=709                                                                                                         \n");
                htmlStr.Append("					style='height: 2.0pt; width: 531pt'>&nbsp;</td>                                                                                                         \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=8 style='mso-height-source: userset; height: 6.0pt'>                                                                                                 \n");
                htmlStr.Append("				<td height=8 class=xl7813844 width=30                                                                                                                       \n");
                htmlStr.Append("					style='height: 6.0pt; width: 20.7pt'></td>                                                                                                              \n");
                htmlStr.Append("				<td colspan=26 class=xl15138441></td>                                                                                                                       \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 19.42pt;'>                                                                                               \n");
                htmlStr.Append("				<td colspan=3 height=24 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 19.42pt; width: 109pt'>&#272;&#417;n v&#7883;                                                                                             \n");
                htmlStr.Append("					bán hàng:</td>                                                                                                                                          \n");
                htmlStr.Append("				<td colspan=23 class=xl6313844 width=564 style='width: 422pt'>" + dt.Rows[0]["SELLER_NAME"] + "</td>                                                                        \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 19.42pt'>                                                                                               \n");
                htmlStr.Append("				<td colspan=3 height=24 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 19.42pt; width: 109pt'>&#272;&#7883;a ch&#7881;:</td>                                                                                     \n");
                htmlStr.Append("				<td colspan=23 class=xl6313844 width=564 style='width: 422pt'>" + dt.Rows[0]["SELLER_ADDRESS"] + "</td>                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 19.42pt'>                                                                                               \n");
                htmlStr.Append("				<td colspan=3 height=24 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 19.42pt; width: 109pt'>S&#7889; tài kho&#7843;n:</td>                                                                                     \n");
                htmlStr.Append("				<td colspan=23 class=xl6313844 width=564 style='width: 422pt'>" + dt.Rows[0]["SELLER_ACCOUNTNO"] + " </td>                                                                  \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr style='mso-height-source: userset; height: 20.00pt'>                                                                                               \n");
                htmlStr.Append("				<td colspan=3 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 20.00pt; width: 109pt'>&#272;i&#7879;n                                                                                                    \n");
                htmlStr.Append("					tho&#7841;i:</td>                                                                                                                                       \n");
                htmlStr.Append("				<td class=xl6313844 style='width: 60.3pt'>" + dt.Rows[0]["SELLER_TEL"] + "</td>                                                                                 \n");
                htmlStr.Append("				<td class=xl6313844 style='width: 35.0pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 style='width: 70pt'>MS:&nbsp;&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("				<td width=23 coslpan =20 >                                                                                                                   \n");
                htmlStr.Append("					<table border=0 cellpadding=0 cellspacing=0 width=350                                                                                                  \n");
                htmlStr.Append("						style='border-collapse: collapse; table-layout: fixed;'>                                                                                            \n");
                htmlStr.Append("						                                                                                                                                                    \n");
                htmlStr.Append("						<tr height=6 style='mso-height-source: userset; height: 17.125pt'>                                                                                    \n");
                htmlStr.Append("							<td class=xl7013844 width=23 style='width: 3.0pt; height: 17.125pt'>" + l_taxcode_0 + "</td>                                                        \n");
                htmlStr.Append("							<td class=xl7013844 width=23 style='border-left: none; width: 3.0pt'>" + l_taxcode_1 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl7013844 width=23 style='border-left: none; width: 3.0pt'>" + l_taxcode_2 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_taxcode_3 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_taxcode_4 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl11913844 width=23                                                                                                                   \n");
                htmlStr.Append("								style='border-right: 1pt solid black; border-left: none; width: 3.0pt'>" + l_taxcode_5 + "</td>                                              \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_taxcode_6 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_taxcode_7 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl11913844 width=26                                                                                                                   \n");
                htmlStr.Append("								style='border-right: 1pt solid black; border-left: none; width: 3.0pt'>" + l_taxcode_8 + "</td>                                              \n");
                htmlStr.Append("							<td class=xl6613844 style='border-left: none; width: 3.0pt'>" + l_taxcode_9 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl7113844 style='border-left: none; width: 1.5pt'></td>                                                                \n");
                htmlStr.Append("							<td class=xl6913844  style='width: 1.5pt'></td>                                                                                         \n");
                htmlStr.Append("							<td class=xl6613844 style='width: 3.0pt'></td>                                                                       \n");
                htmlStr.Append("							<td class=xl6613844 style='border-left: none; width: 3.0pt'></td>                                                    \n");
                htmlStr.Append("							<td class=xl6613844 style='border-left: none; width: 3.0pt'></td>                                                    \n");
                htmlStr.Append("							<td class=xl6613844 style='border-left: none; width: 3pt'></td>                                                    \n");
                htmlStr.Append("							<td class=xl7113844 style='border-left: none; width: 3pt'></td>                                                                      \n");
                htmlStr.Append("						</tr>                                                                                                                                               \n");
                htmlStr.Append("					</table>                                                                                                                                                \n");
                htmlStr.Append("				                                                                                                                                                            \n");
                htmlStr.Append("				</td>                                                                                                                                                       \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 19.42pt'>                                                                                               \n");
                htmlStr.Append("				<td colspan=3 height=24 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 19.42pt; width: 109pt'>Fax:</td>                                                                                                          \n");
                htmlStr.Append("				<td class=xl6313844 width=90 style='width: 60.3pt'>" + dt.Rows[0]["SELLER_FAX"] + "</td>                                                                                 \n");
                htmlStr.Append("				<td class=xl6313844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=8 style='mso-height-source: userset; height: 7.0pt'>                                                                                                 \n");
                htmlStr.Append("				<td height=8 class=xl6313844 width=30                                                                                                                       \n");
                htmlStr.Append("					style='height: 7.0pt; width: 20.7pt'></td>                                                                                                              \n");
                htmlStr.Append("				<td class=xl6313844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=16 style='width: 10.8pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=8 style='mso-height-source: userset; height: 2.0pt'>                                                                                                 \n");
                htmlStr.Append("				<td colspan=26 height=8 class=xl12113844 width=709                                                                                                          \n");
                htmlStr.Append("					style='height: 2.0pt; width: 531pt'>&nbsp;</td>                                                                                                         \n");
                htmlStr.Append("				<td class=xl15138441></td>                                                                                                                                  \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=2 style='mso-height-source: userset; height: 1.5pt'>                                                                                                 \n");
                htmlStr.Append("				<td colspan=27 height=2 class=xl15138441 style='height: 1.5pt'></td>                                                                                        \n");
                htmlStr.Append("                                                                                                                                                                            \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=25 style='mso-height-source: userset; height: 18.20pt'>                                                                                             \n");
                htmlStr.Append("				<td colspan=3 height=25 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 18.20pt; width: 109pt'>H&#7885; tên                                                                                                     \n");
                htmlStr.Append("					ng&#432;&#7901;i mua hàng:</td>                                                                                                                         \n");
                htmlStr.Append("				<td colspan=23 class=xl6313844 width=564 style='width: 422pt'>" + dt.Rows[0]["BUYER"] + "</td>                                                                       \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=25 style='mso-height-source: userset; height: 18.50pt'>                                                                                             \n");
                htmlStr.Append("				<td colspan=3 height=25 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 18.50pt; width: 109pt'>Tên &#273;&#417;n                                                                                                \n");
                htmlStr.Append("					v&#7883;:</td>                                                                                                                                          \n");
                htmlStr.Append("				<td colspan=23 class=xl6313844 width=564 style='width: 422pt'>" + dt.Rows[0]["BUYERLEGALNAME"] + "</td>                                                                        \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=25 style='mso-height-source: userset; height: 19.00pt'>                                                                                             \n");
                htmlStr.Append("				<td colspan=3 height=25 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 19.00pt; width: 109pt'>&#272;&#7883;a                                                                                                   \n");
                htmlStr.Append("					ch&#7881;:</td>                                                                                                                                         \n");
                htmlStr.Append("				<td colspan=23 class=xl11813844 width=564 style='width: 422pt'>" + dt.Rows[0]["BUYERADDRESS"] + "</td>                                                                          \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=25 style='mso-height-source: userset; height: 19.00pt'>                                                                                             \n");
                htmlStr.Append("				<td colspan=3 height=25 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 19.00pt; width: 109pt'>S&#7889; tài                                                                                                     \n");
                htmlStr.Append("					kho&#7843;n:</td>                                                                                                                                       \n");
                htmlStr.Append("				<td colspan=23 class=xl11813844 width=564 style='width: 422pt'></td>                                                                                        \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=25 style='mso-height-source: userset; height: 20.00pt'>                                                                                             \n");
                htmlStr.Append("				<td colspan=3 class=xl6313844 width=145                                                                                                           \n");
                htmlStr.Append("					style='height: 20.00pt; width: 109pt'>Hình th&#7913;c thanh                                                                                            \n");
                htmlStr.Append("					toán:</td>                                                                                                                                              \n");
                htmlStr.Append("				<td colspan=2 class=xl11813844 style='width: 100pt'>" + dt.Rows[0]["PAYMENTMETHODCK"] + "</td>                                                                    \n");
                htmlStr.Append("				<td class=xl6313844 style='width: 80.00pt;'>MS:&nbsp;&nbsp;</td>                                                                                                 \n");
                htmlStr.Append("				<td coslpan =20 >                                                                                                                                  \n");
                htmlStr.Append("					<table border=0 cellpadding=0 cellspacing=0 width=350                                                                                                   \n");
                htmlStr.Append("						style='border-collapse: collapse; table-layout: fixed;'>                                                                                            \n");
                htmlStr.Append("						                                                                                                                                                    \n");
                htmlStr.Append("						<tr height=6 style='mso-height-source: userset; height: 17.145pt'>                                                                                    \n");
                htmlStr.Append("							<td class=xl7013844 width=23 style='width: 3.0pt; height: 17.145pt'>" + l_buyertax_0 + "</td>                                                        \n");
                htmlStr.Append("							<td class=xl7013844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_1 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl7013844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_2 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_3 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_4 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl11913844 width=23                                                                                                                   \n");
                htmlStr.Append("								style='border-right: 1pt solid black; border-left: none; width: 3.0pt'>" + l_buyertax_5 + "</td>                                              \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_6 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_7 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl11913844 width=26                                                                                                                   \n");
                htmlStr.Append("								style='border-right: 1pt solid black; border-left: none; width: 3.0pt'>" + l_buyertax_8 + "</td>                                              \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_9 + "</td>                                                     \n");
                htmlStr.Append("							<td class=xl7113844 width=13 style='border-left: none; width: 1.5pt'></td>                                                                \n");
                htmlStr.Append("							<td class=xl6913844 width=13 style='width: 1.5pt'></td>                                                                                         \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='width: 3.0pt'>" + l_buyertax_10 + "</td>                                                                       \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_11 + "</td>                                                    \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'>" + l_buyertax_12 + "</td>                                                    \n");
                htmlStr.Append("							<td class=xl6613844 width=23 style='border-left: none; width: 3.0pt'></td>                                                    \n");
                htmlStr.Append("							<td class=xl7113844 width=26 style='border-left: none; width: 3pt'></td>                                                                      \n");
                htmlStr.Append("						</tr>                                                                                                                                               \n");
                htmlStr.Append("					</table>                                                                                                                                                \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=13 style='mso-height-source: userset; height: 12.50pt'>                                                                                              \n");
                htmlStr.Append("				<td height=13 class=xl1513844 style='height: 12.50pt'></td>                                                                                                 \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");
                htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 18.00pt'>                                                                                               \n");
                htmlStr.Append("				<td height=24 class=xl8013844 width=30                                                                                                                      \n");
                htmlStr.Append("					style='height: 18.00pt; width: 20.7pt'>STT</td>                                                                                                          \n");
                htmlStr.Append("				<td colspan=2 class=xl8013844 width=115                                                                                                                     \n");
                htmlStr.Append("					style='border-left: none; width: 86pt'>MÃ SP</td>                                                                                                       \n");
                htmlStr.Append("				<td colspan=5 class=xl8013844 width=233                                                                                                                     \n");
                htmlStr.Append("					style='border-left: none; width: 174pt'>TÊN SP</td>                                                                                                     \n");
                htmlStr.Append("				<td colspan=2 class=xl8013844 width=46                                                                                                                      \n");
                htmlStr.Append("					style='border-left: none; width: 34pt'>KH&#7892;</td>                                                                                                   \n");
                htmlStr.Append("				<td colspan=2 class=xl8013844 width=36                                                                                                                      \n");
                htmlStr.Append("					style='border-left: none; width: 27pt'>&#272;VT</td>                                                                                                    \n");
                htmlStr.Append("				<td colspan=4 class=xl8013844 width=69                                                                                                                      \n");
                htmlStr.Append("					style='border-left: none; width: 52pt'>S&#7888;                                                                                                         \n");
                htmlStr.Append("					L&#431;&#7906;NG</td>                                                                                                                                   \n");
                htmlStr.Append("				<td colspan=5 class=xl8013844 width=75                                                                                                                      \n");
                htmlStr.Append("					style='border-left: none; width: 57pt'>&#272;&#416;N GIÁ</td>                                                                                           \n");
                htmlStr.Append("				<td colspan=5 class=xl8013844 width=105                                                                                                                     \n");
                htmlStr.Append("					style='border-left: none; width: 78pt'>THÀNH TI&#7872;N</td>                                                                                            \n");
                htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
                htmlStr.Append("			</tr>                                                                                                                                                           \n");


                v_rowHeight = "30.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 24;

                v_rowHeightLast = "28.0pt";// "23.5pt";
                v_rowHeightLastNumber = 26;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "30.0pt"; //"26.5pt";    
                        v_rowHeightLast = "28.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 24;//27.5;
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
                        htmlStr.Append("					<tr height=26 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                      \n");
                        htmlStr.Append("						<td height=26 class=xl7913844 width=30                                                                                                              \n");
                        htmlStr.Append("							style='height: 18.09pt; border-top: none; width: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; font-size:9pt;'>" + dt_d.Rows[0]["STT"] + "</td>                                                                     \n");
                        htmlStr.Append("						<td colspan=2 class=xl11513844 width=115                                                                                                            \n");
                        htmlStr.Append("							style='border-left: none; width: 86pt; font-size:9pt;'>&nbsp; " + dt_d.Rows[0]["ITEM_CODE"] + "</td>                                                                                 \n");
                        htmlStr.Append("						<td colspan=5 class=xl11713844 width=233                                                                                                            \n");
                        htmlStr.Append("							style='border-left: none; width: 174pt'>&nbsp;" + dt_d.Rows[0]["ITEM_NAME"] + "</td>                                                                                \n");
                        htmlStr.Append("						<td colspan=2 class=xl7913844 width=46                                                                                                              \n");
                        htmlStr.Append("							style='border-left: none; width: 34pt; font-size:9pt;'>" + (dt_d.Rows[0]["ITEM_CODE"].ToString() == "Y" ? "" : dt_d.Rows[0]["ATTRIBUTE_3"].ToString()) + "</td>                                                                                       \n");
                        htmlStr.Append("						<td colspan=2 class=xl7913844 width=36                                                                                                              \n");
                        htmlStr.Append("							style='border-left: none; width: 27pt; font-size:9pt;'>" + dt_d.Rows[0]["ITEM_UOM"].ToString() + "</td>                                                                                       \n");
                        htmlStr.Append("						<td colspan=4 class=xl11413844 width=69                                                                                                             \n");
                        htmlStr.Append("							style='border-left: none; width: 52pt; font-size:9pt;'>" + dt_d.Rows[0]["QTY"] + "&nbsp;</td>                                                                                 \n");
                        htmlStr.Append("						<td colspan=5 class=xl11613844 width=75                                                                                                             \n");
                        htmlStr.Append("							style='border-left: none; width: 57pt; font-size:9pt;'>" + dt_d.Rows[0]["U_PRICE"] + "&nbsp;</td>                                                                                 \n");
                        htmlStr.Append("						<td colspan=5 class=xl11613844 width=105                                                                                                            \n");
                        htmlStr.Append("							style='border-left: none; width: 78pt; font-size:9pt;'>" + dt_d.Rows[0]["NET_TR_AMT"] + "&nbsp;</td>                                                                                 \n");
                        htmlStr.Append("						<td class=xl1513844></td>                                                                                                                           \n");
                        htmlStr.Append("					</tr>                                                                                                                                                   \n");

                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("					<tr height=26 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                      \n");
                            htmlStr.Append("						<td height=26 class=xl7913844 width=30                                                                                                              \n");
                            htmlStr.Append("							style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none; width: 20.7pt; font-size:9pt;'>" + dt_d.Rows[dtR]["STT"] + "</td>                                                                     \n");
                            htmlStr.Append("						<td colspan=2 class=xl11513844 width=115                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 86pt; font-size: 9pt;'>&nbsp; " + dt_d.Rows[dtR]["ITEM_CODE"] + "</td>                                                                                 \n");
                            htmlStr.Append("						<td colspan=5 class=xl11713844 width=233                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 174pt'>&nbsp;" + dt_d.Rows[dtR]["ITEM_NAME"] + "</td>                                                                                \n");
                            htmlStr.Append("						<td colspan=2 class=xl7913844 width=46                                                                                                              \n");
                            htmlStr.Append("							style='border-left: none; width: 34pt; font-size: 9pt;'>" + (dt_d.Rows[dtR]["ITEM_CODE"].ToString() == "Y" ? "" : dt_d.Rows[dtR]["ATTRIBUTE_3"].ToString()) + "</td>                                                                                       \n");
                            htmlStr.Append("						<td colspan=2 class=xl7913844 width=36                                                                                                              \n");
                            htmlStr.Append("							style='border-left: none; width: 27pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["ITEM_UOM"].ToString() + "</td>                                                                                       \n");
                            htmlStr.Append("						<td colspan=4 class=xl11413844 width=69                                                                                                             \n");
                            htmlStr.Append("							style='border-left: none; width: 52pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["QTY"] + "&nbsp;</td>                                                                                 \n");
                            htmlStr.Append("						<td colspan=5 class=xl11413844 width=75                                                                                                             \n");
                            htmlStr.Append("							style='border-left: none; width: 57pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["U_PRICE"] + "&nbsp;</td>                                                                                 \n");
                            htmlStr.Append("						<td colspan=5 class=xl11413844 width=105                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 78pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["NET_TR_AMT"] + "&nbsp;</td>                                                                                 \n");
                            htmlStr.Append("						<td class=xl1513844></td>                                                                                                                           \n");
                            htmlStr.Append("					</tr>                                                                                                                                                   \n");

                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("					<tr height=26 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                      \n");
                                htmlStr.Append("						<td height=26 class=xl7913844 width=30                                                                                                              \n");
                                htmlStr.Append("							style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none; width: 20.7pt; font-size:9pt;'>" + dt_d.Rows[dtR]["STT"] + "</td>                                                                     \n");
                                htmlStr.Append("						<td colspan=2 class=xl11513844 width=115                                                                                                            \n");
                                htmlStr.Append("							style='border-left: none; width: 86pt; font-size: 9pt;'>&nbsp; " + dt_d.Rows[dtR]["ITEM_CODE"] + "</td>                                                                                 \n");
                                htmlStr.Append("						<td colspan=5 class=xl11713844 width=233                                                                                                            \n");
                                htmlStr.Append("							style='border-left: none; width: 174pt'>&nbsp;" + dt_d.Rows[dtR]["ITEM_NAME"] + "</td>                                                                                \n");
                                htmlStr.Append("						<td colspan=2 class=xl7913844 width=46                                                                                                              \n");
                                htmlStr.Append("							style='border-left: none; width: 34pt; font-size: 9pt;'>" + (dt_d.Rows[dtR]["ITEM_CODE"].ToString() == "Y" ? "" : dt_d.Rows[dtR]["ATTRIBUTE_3"].ToString()) + "</td>                                                                                       \n");
                                htmlStr.Append("						<td colspan=2 class=xl7913844 width=36                                                                                                              \n");
                                htmlStr.Append("							style='border-left: none; width: 27pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["ITEM_UOM"].ToString() + "</td>                                                                                       \n");
                                htmlStr.Append("						<td colspan=4 class=xl11413844 width=69                                                                                                             \n");
                                htmlStr.Append("							style='border-left: none; width: 52pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["QTY"] + "&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("						<td colspan=5 class=xl11413844 width=75                                                                                                             \n");
                                htmlStr.Append("							style='border-left: none; width: 57pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["U_PRICE"] + "&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("						<td colspan=5 class=xl11413844 width=105                                                                                                            \n");
                                htmlStr.Append("							style='border-left: none; width: 78pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["NET_TR_AMT"] + "&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("						<td class=xl1513844></td>                                                                                                                           \n");
                                htmlStr.Append("					</tr>                                                                                                                                                   \n");
                            }
                            else
                            {
                                htmlStr.Append("					<tr height=26 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                      \n");
                                htmlStr.Append("						<td height=26 class=xl7913844 width=30                                                                                                              \n");
                                htmlStr.Append("							style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none; width: 20.7pt; font-size:9pt;'>" + dt_d.Rows[dtR]["STT"] + "</td>                                                                     \n");
                                htmlStr.Append("						<td colspan=2 class=xl11513844 width=115                                                                                                            \n");
                                htmlStr.Append("							style='border-left: none; width: 86pt; font-size: 9pt;'>&nbsp; " + dt_d.Rows[dtR]["ITEM_CODE"] + "</td>                                                                                 \n");
                                htmlStr.Append("						<td colspan=5 class=xl11713844 width=233                                                                                                            \n");
                                htmlStr.Append("							style='border-left: none; width: 174pt'>&nbsp;" + dt_d.Rows[dtR]["ITEM_NAME"] + "</td>                                                                                \n");
                                htmlStr.Append("						<td colspan=2 class=xl7913844 width=46                                                                                                              \n");
                                htmlStr.Append("							style='border-left: none; width: 34pt; font-size: 9pt;'>" + (dt_d.Rows[dtR]["ITEM_CODE"].ToString() == "Y" ? "" : dt_d.Rows[dtR]["ATTRIBUTE_3"].ToString()) + "</td>                                                                                       \n");
                                htmlStr.Append("						<td colspan=2 class=xl7913844 width=36                                                                                                              \n");
                                htmlStr.Append("							style='border-left: none; width: 27pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["ITEM_UOM"].ToString() + "</td>                                                                                       \n");
                                htmlStr.Append("						<td colspan=4 class=xl11413844 width=69                                                                                                             \n");
                                htmlStr.Append("							style='border-left: none; width: 52pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["QTY"] + "&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("						<td colspan=5 class=xl11413844 width=75                                                                                                             \n");
                                htmlStr.Append("							style='border-left: none; width: 57pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["U_PRICE"] + "&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("						<td colspan=5 class=xl11413844 width=105                                                                                                            \n");
                                htmlStr.Append("							style='border-left: none; width: 78pt; font-size: 9pt;'>" + dt_d.Rows[dtR]["NET_TR_AMT"] + "&nbsp;</td>                                                                                 \n");
                                htmlStr.Append("						<td class=xl1513844></td>                                                                                                                           \n");
                                htmlStr.Append("					</tr>                                                                                                                                                   \n");
                            }

                        }
                    }
                    else
                    { // dong giua
                      // 


                        htmlStr.Append("				<tr height=26 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>                                                                                          \n");
                        htmlStr.Append("					<td height=26 class=xl7913844 width=30                                                                                                                  \n");
                        htmlStr.Append("						style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-top: none; width: 20.7pt;font-size:9.0pt;'>" + dt_d.Rows[dtR]["STT"] + "</td>                                                                         \n");
                        htmlStr.Append("					<td colspan=2 class=xl11713844 width=115                                                                                                                \n");
                        htmlStr.Append("						style='border-left: none; width: 86pt; font-size: 9pt;'>&nbsp; " + dt_d.Rows[dtR]["ITEM_CODE"] + "</td>                                                                                     \n");
                        htmlStr.Append("					<td colspan=5 class=xl11713844 width=233                                                                                                                \n");
                        htmlStr.Append("						style='border-left: none; width: 174pt'>&nbsp;" + dt_d.Rows[dtR]["ITEM_NAME"] + "</td>                                                                                    \n");
                        htmlStr.Append("					<td colspan=2 class=xl7913844 width=46                                                                                                                  \n");
                        htmlStr.Append("						style='border-left: none; width: 34pt; font-size:9.0pt;'>" + (dt_d.Rows[dtR]["ITEM_CODE"].ToString() == "Y" ? "" : dt_d.Rows[dtR]["ATTRIBUTE_3"].ToString()) + "</td>                                                                                           \n");
                        htmlStr.Append("					<td colspan=2 class=xl7913844 width=36                                                                                                                  \n");
                        htmlStr.Append("						style='border-left: none; width: 27pt; font-size:9.0pt;'>" + dt_d.Rows[dtR]["ITEM_UOM"].ToString() + "</td>                                                                                           \n");
                        htmlStr.Append("					<td colspan=4 class=xl11413844 width=69                                                                                                                 \n");
                        htmlStr.Append("						style='border-left: none; width: 52pt; font-size:9.0pt;'>" + dt_d.Rows[dtR]["QTY"] + "&nbsp;</td>                                                                                     \n");
                        htmlStr.Append("					<td colspan=5 class=xl11613844 width=75                                                                                                                 \n");
                        htmlStr.Append("						style='border-left: none; width: 57pt; font-size:9.0pt;'>" + dt_d.Rows[dtR]["U_PRICE"] + "&nbsp;</td>                                                                                     \n");
                        htmlStr.Append("					<td colspan=5 class=xl11613844 width=105                                                                                                                \n");
                        htmlStr.Append("						style='border-left: none; width: 78pt; font-size:9.0pt;'>" + dt_d.Rows[dtR]["NET_TR_AMT"] + "&nbsp;</td>                                                                                     \n");
                        htmlStr.Append("					<td class=xl1513844></td>                                                                                                                               \n");
                        htmlStr.Append("				</tr>                                                                                                                                                       \n");
                        htmlStr.Append("			                                                                                                                                                                \n");


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
                    v_spacePerPage = 18;
                }

                if (k == v_countNumberOfPages - 1 && page[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page[k]; i++)
                    {
                        if (i == (rowsPerPage - page[k] - 1))
                        {
                            htmlStr.Append("					<tr height=26 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "';'>                                                                                      \n");
                            htmlStr.Append("						<td height=26 class=xl7913844 width=30                                                                                                              \n");
                            htmlStr.Append("							style='height: " + v_rowHeightEmptyLast + "'; border-top: none; width: 20.7pt;font-size:9.0pt;'>&nbsp;</td>                                                                            \n");
                            htmlStr.Append("						<td colspan=2 class=xl11513844 width=115                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 86pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=5 class=xl11713844 width=233                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 174pt'>&nbsp;</td>                                                                                             \n");
                            htmlStr.Append("						<td colspan=2 class=xl7913844 width=46                                                                                                              \n");
                            htmlStr.Append("							style='border-left: none; width: 34pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=2 class=xl7913844 width=36                                                                                                              \n");
                            htmlStr.Append("							style='border-left: none; width: 27pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=4 class=xl11413844 width=69                                                                                                             \n");
                            htmlStr.Append("							style='border-left: none; width: 52pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=5 class=xl11413844 width=75                                                                                                             \n");
                            htmlStr.Append("							style='border-left: none; width: 57pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=5 class=xl11413844 width=105                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 78pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td class=xl1513844></td>                                                                                                                           \n");
                            htmlStr.Append("					</tr>                                                                                                                                                   \n");



                        }
                        else
                        {
                            htmlStr.Append("					<tr height=26 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + ";'>                                                                                      \n");
                            htmlStr.Append("						<td height=26 class=xl7913844 width=30                                                                                                              \n");
                            htmlStr.Append("							style='height: " + v_rowHeightEmptyLast + "; border-top: none; width: 20.7pt'>&nbsp;</td>                                                                            \n");
                            htmlStr.Append("						<td colspan=2 class=xl11513844 width=115                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 86pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=5 class=xl11713844 width=233                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 174pt'>&nbsp;</td>                                                                                             \n");
                            htmlStr.Append("						<td colspan=2 class=xl7913844 width=46                                                                                                              \n");
                            htmlStr.Append("							style='border-left: none; width: 34pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=2 class=xl7913844 width=36                                                                                                              \n");
                            htmlStr.Append("							style='border-left: none; width: 27pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=4 class=xl11413844 width=69                                                                                                             \n");
                            htmlStr.Append("							style='border-left: none; width: 52pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=5 class=xl11613844 width=75                                                                                                             \n");
                            htmlStr.Append("							style='border-left: none; width: 57pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td colspan=5 class=xl11613844 width=105                                                                                                            \n");
                            htmlStr.Append("							style='border-left: none; width: 78pt'>&nbsp;</td>                                                                                              \n");
                            htmlStr.Append("						<td class=xl1513844></td>                                                                                                                           \n");
                            htmlStr.Append("					</tr>                                                                                                                                                   \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {

                    htmlStr.Append("					<tr height=26 style='mso-height-source: userset; height: " + (v_spacePerPage).ToString() + "pt';'>                                                                                      \n");
                    htmlStr.Append("						<td height=26 class=xl7913845 width=30                                                                                                              \n");
                    htmlStr.Append("							style='height: " + (v_spacePerPage).ToString() + "pt';border-left: 1pt solid black;border-bottom: 1pt solid black;border-top: none;border-right: none; width: 20.7pt;font-size:9.0pt;'>&nbsp;</td>                                                                            \n");
                    htmlStr.Append("						<td colspan=2 class=xl11513844 width=115                                                                                                            \n");
                    htmlStr.Append("							style='border-left: none;border-right: none; width: 86pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td colspan=5 class=xl11713844 width=233                                                                                                            \n");
                    htmlStr.Append("							style='border-left: none;border-right: none; width: 174pt'>&nbsp;</td>                                                                                             \n");
                    htmlStr.Append("						<td colspan=2 class=xl7913844 width=46                                                                                                              \n");
                    htmlStr.Append("							style='border-left: none;border-right: none; width: 34pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td colspan=2 class=xl7913844 width=36                                                                                                              \n");
                    htmlStr.Append("							style='border-left: none;border-right: none; width: 27pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td colspan=4 class=xl11413844 width=69                                                                                                             \n");
                    htmlStr.Append("							style='border-left: none;border-right: none; width: 52pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td colspan=5 class=xl11413844 width=75                                                                                                             \n");
                    htmlStr.Append("							style='border-left: none;border-right: none; width: 57pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td colspan=5 class=xl11413844 width=105                                                                                                            \n");
                    htmlStr.Append("							style='border-left: none; width: 78pt'>&nbsp;</td>                                                                                              \n");
                    htmlStr.Append("						<td class=xl1513844></td>                                                                                                                           \n");
                    htmlStr.Append("					</tr>                                                                                                                                                   \n");
                    htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    htmlStr.Append("		<tr height=5  style='height: 18pt'>                                                                                                                                                                \n");

                    htmlStr.Append("		</tr>      																																														\n");
                    htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             
            htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 19.25pt'>                                                                                               \n");
            htmlStr.Append("				<td colspan=4 height=24 class=xl10413844 width=235                                                                                                          \n");
            htmlStr.Append("					style='height: 19.25pt; width: 176pt'>T&#7927; giá (Exchange                                                                                             \n");
            htmlStr.Append("					rate):<span style='mso-spacerun: yes'> </span>                                                                                                          \n");
            htmlStr.Append("				</td>                                                                                                                                                       \n");
            htmlStr.Append("				<td colspan=2 class=xl7313844 width=97 style='width: 73pt'>" + dt.Rows[0]["EXCHANGERATE_NO"] + "</td>                                                                      \n");
            htmlStr.Append("				<td class=xl7313844 width=23 style='border-top: none; width: 15.3pt'>&nbsp;</td>                                                                            \n");
            htmlStr.Append("				<td class=xl7313844 width=23 style='border-top: none; width: 15.3pt'>&nbsp;</td>                                                                            \n");
            htmlStr.Append("				<td colspan=13 class=xl10413844 width=226                                                                                                                   \n");
            htmlStr.Append("					style='border-right: 1pt solid black; width: 170pt'>C&#7897;ng                                                                                         \n");
            htmlStr.Append("					ti&#7873;n hàng:&nbsp;</td>                                                                                                                                   \n");
            htmlStr.Append("				<td colspan=5 class=xl10713844 width=105                                                                                                                    \n");
            htmlStr.Append("					style='border-right: 1pt solid black; border-left: none; width: 78pt'>" + amount_net + "&nbsp;</td>                                                        \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 19.25pt'>                                                                                               \n");
            htmlStr.Append("				<td colspan=4 height=24 class=xl11013844 width=235                                                                                                          \n");
            htmlStr.Append("					style='height: 19.25pt; width: 176pt'>Thu&#7871; su&#7845;t                                                                                              \n");
            htmlStr.Append("					GTGT:<span style='mso-spacerun: yes'> </span>                                                                                                           \n");
            htmlStr.Append("				</td>                                                                                                                                                       \n");
            htmlStr.Append("				<td colspan=2 class=xl11213844 width=97 style='width: 73pt'>" + dt.Rows[0]["taxrate"] + "</td>                                                                            \n");
            htmlStr.Append("				<td class=xl6413844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6413844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td colspan=13 class=xl11013844 width=226                                                                                                                   \n");
            htmlStr.Append("					style='border-right: 1pt solid black; width: 170pt'>Ti&#7873;n                                                                                         \n");
            htmlStr.Append("					thu&#7871; GTGT:&nbsp;</td>                                                                                                                                   \n");
            htmlStr.Append("				<td colspan=5 class=xl10713844 width=105                                                                                                                    \n");
            htmlStr.Append("					style='border-right: 1pt solid black; border-left: none; width: 78pt'>" + amount_vat + "&nbsp;</td>                                                      \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 19.25pt'>                                                                                               \n");
            htmlStr.Append("				<td colspan=4 height=24 class=xl9613844 width=235                                                                                                           \n");
            htmlStr.Append("					style='height: 19.25pt; width: 176pt'>T&#7893;ng c&#7897;ng USD                                                                                          \n");
            htmlStr.Append("					(Amount USD):<span style='mso-spacerun: yes'> </span>                                                                                                   \n");
            htmlStr.Append("				</td>                                                                                                                                                       \n");
            htmlStr.Append("				<td colspan=2 class=xl7413844 width=97 style='width: 73pt'>" + amount_trans + "&nbsp;</td>                                                                       \n");
            htmlStr.Append("				<td class=xl7413844 width=23 style='width: 15.3pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("				<td class=xl7413844 width=23 style='width: 15.3pt'>&nbsp;</td>                                                                                              \n");
            htmlStr.Append("				<td colspan=13 class=xl9613844 width=226                                                                                                                    \n");
            htmlStr.Append("					style='border-right: 1pt solid black; width: 170pt'>T&#7893;ng                                                                                         \n");
            htmlStr.Append("					c&#7897;ng ti&#7873;n thanh toán:&nbsp;</td>                                                                                                                  \n");
            htmlStr.Append("				<td colspan=5 class=xl9913844 width=105                                                                                                                     \n");
            htmlStr.Append("					style='border-right: 1pt solid black; border-left: none; width: 78pt'>" + amount_total + "&nbsp;</td>                                               \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=6 style='mso-height-source: userset; height: 3.25pt'>                                                                                                \n");
            htmlStr.Append("				<td height=6 colspan=26 class=xl7513844 width=30                                                                                                            \n");
            htmlStr.Append("					style='height: 3.25pt; border-top: none; width: 20.7pt'>&nbsp;</td>                                                                                     \n");
            htmlStr.Append("				                                                                                                                                                            \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=24 style='mso-height-source: userset; height: 18.00pt'>                                                                                               \n");
            htmlStr.Append("				<td colspan=27 height=24 class=xl10213844 width=709                                                                                                         \n");
            htmlStr.Append("					style='height: 18.00pt; width: 531pt'>S&#7889; ti&#7873;n                                                                                                \n");
            htmlStr.Append("					(vi&#7871;t b&#7857;ng ch&#7919;):<span style='mso-spacerun: yes'>  " + read_prive + "</span>                                                             \n");
            htmlStr.Append("				</td>                                                                                                                                                       \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=6 style='mso-height-source: userset; height: 10.25pt'>                                                                                               \n");
            htmlStr.Append("				<td height=6 colspan=26 class=xl7513844 width=30                                                                                                            \n");
            htmlStr.Append("					style='height: 10.25pt; border-top: none; width: 20.7pt'>&nbsp;</td>                                                                                    \n");
            htmlStr.Append("				                                                                                                                                                            \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=18 style='mso-height-source: userset; height: 19.25pt'>                                                                                             \n");
            htmlStr.Append("				<td colspan=4 height=18 class=xl6513844 width=235                                                                                                           \n");
            htmlStr.Append("					style='height: 12.825pt; width: 176pt'>NG&#431;&#7900;I MUA HÀNG                                                                                                 \n");
            htmlStr.Append("					</td>                                                                                                                        \n");
            htmlStr.Append("				<td colspan=4 class=xl6513844 width=143 style='width: 107pt'>                                                                               \n");
            htmlStr.Append("					</td>                                                                                                                                           \n");
            htmlStr.Append("				<td class=xl8113844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl8113844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td colspan=16 class=xl10313844>NG&#431;&#7900;I BÁN HÀNG</td>                                                                                              \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=18 style='mso-height-source: userset; height: 20.00pt'>                                                                                             \n");
            htmlStr.Append("				<td colspan=4 height=18 class=xl6513844 width=235                                                                                                           \n");
            htmlStr.Append("					style='height: 20.00pt; width: 176pt'>(Ký, h&#7885; tên)</td>                                                                                          \n");
            htmlStr.Append("				<td colspan=4 class=xl6513844 width=143 style='width: 107pt'>                                                                                           \n");
            htmlStr.Append("					</td>                                                                                                                                      \n");
            htmlStr.Append("				<td class=xl8113844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl8113844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td colspan=16 class=xl6513844 width=285 style='width: 214pt'>(Ký,                                                                                          \n");
            htmlStr.Append("					h&#7885; tên)</td>                                                                                                                                      \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=27 style='mso-height-source: userset; height: 37.05pt'>                                                                                              \n");
            htmlStr.Append("				<td height=27 class=xl6513844 width=30                                                                                                                      \n");
            htmlStr.Append("					style='height: 37.05pt; width: 20.7pt'></td>                                                                                                            \n");
            htmlStr.Append("				<td class=xl6513844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=16 style='width: 10.8pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=13 style='width: 9pt'></td>                                                                                                       \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl6513844 width=13 style='width: 9pt'></td>                                                                                                       \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=13 style='width: 9pt'></td>                                                                                                       \n");
            htmlStr.Append("				<td class=xl6513844 width=13 style='width: 9pt'></td>                                                                                                       \n");
            htmlStr.Append("				<td class=xl6513844 width=13 style='width: 9pt'></td>                                                                                                       \n");
            htmlStr.Append("				<td class=xl6513844 width=10 style='width: 7.2pt'></td>                                                                                                     \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=26 style='width: 17.1pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=21 style='mso-height-source: userset; height: 20.00pt'>                                                                                             \n");
            htmlStr.Append("				<td height=21 class=xl6513844 width=30                                                                                                                      \n");
            htmlStr.Append("					style='height: 20.00pt; width: 20.7pt'></td>                                                                                                           \n");
            htmlStr.Append("				<td class=xl6513844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=16 style='width: 10.8pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");

            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("				<td colspan=16 height=21 class=xl8913844 width=285                                                                                                          \n");
                htmlStr.Append("					style='border-right: 1pt solid black; height: 21.00pt; width: 214pt'                                                                                  \n");
                htmlStr.Append("					align=left valign=top><![if !vml]><span                                                                                                                 \n");
                htmlStr.Append("					style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 148px; margin-top: 8px; width: 81.25px; height: 53.75px'><img                       \n");
                htmlStr.Append("						width=65 height=43                                                                                                                                  \n");
                htmlStr.Append("						src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png'                                                                                \n");
                htmlStr.Append("						v:shapes='Picture_x0020_8'></span>                                                                                                                  \n");
                htmlStr.Append("				<![endif]><span style='mso-ignore: vglayout2'>                                                                                                              \n");
                htmlStr.Append("						<table cellpadding=0 cellspacing=0>                                                                                                                 \n");
                htmlStr.Append("							<tr>                                                                                                                                            \n");
                htmlStr.Append("								<td colspan=16 height=21  width=285                                                                                                         \n");
                htmlStr.Append("									style=' height: 18.05pt; width: 267.50pt'>Signature                                                                                       \n");
                htmlStr.Append("									Valid</td>                                                                                                                              \n");
                htmlStr.Append("							</tr>                                                                                                                                           \n");
                htmlStr.Append("						</table>                                                                                                                                            \n");
                htmlStr.Append("				</span></td>                                                                                                                                                \n");

            }
            else
            {

                htmlStr.Append("				<td colspan=16 height=21 class=xl8913844 width=285                                                                                                          \n");
                htmlStr.Append("					style='border-right: 1pt solid black; height: 21.00pt; width: 214pt; font-size:11pt; '                                                                                  \n");
                htmlStr.Append("					align=left valign=top>&nbsp;Signature Valid</td>                                                                                                              \n");

            }

            htmlStr.Append("				                                                                                                                                                            \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=21 style='mso-height-source: userset; height: 18.00pt'>                                                                                             \n");
            htmlStr.Append("				<td height=21 class=xl6513844 width=30                                                                                                                      \n");
            htmlStr.Append("					style='height: 18.00pt; width: 20.7pt'></td>                                                                                                           \n");
            htmlStr.Append("				<td class=xl6513844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=16 style='width: 10.8pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td colspan=16 class=xl9213844                                                                                                                              \n");
            htmlStr.Append("					style='border-right: 1pt solid black'>&nbsp;<font                                                                                                            \n");
            htmlStr.Append("					class='font713844'>&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                                                       \n");
            htmlStr.Append("					class='font613844'> </font><font class='font813844'>" + dt.Rows[0]["SignedBy"] + "</font></td>                                                                       \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=21 style='mso-height-source: userset; height: 18.00pt'>                                                                                             \n");
            htmlStr.Append("				<td height=21 class=xl8613844 style='height: 18.00pt'></td>                                                                                                \n");
            htmlStr.Append("				<td class=xl8113844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=16 style='width: 10.8pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl8213844 colspan=8>&nbsp;Ngày ký: <font class='font913844'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                                 \n");
            htmlStr.Append("				<td class=xl8313844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8313844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8313844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8313844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8313844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8313844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8313844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8413844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=21 style='mso-height-source: userset; height: 18.00pt'>                                                                                             \n");
            htmlStr.Append("				<td height=21 class=xl8613844 style='height: 18.00pt'></td>                                                                                                \n");
            htmlStr.Append("				<td class=xl6513844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=16 style='width: 10.8pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=21 style='mso-height-source: userset; height: 18.00pt'>                                                                                             \n");
            htmlStr.Append("				<td height=21 class=xl8713844 style='height: 18.00pt'>Mã CQT:</td>                                                                                                                      \n");
            htmlStr.Append("				<td class=xl6513844 width=99 style='width: 66.6pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class='xl6513844 font1113845' width=16 style='width: 10.8pt'><span class='font1113845'>" + dt.Rows[0]["cqt_mccqt_id"] + "</span></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=90 style='width: 60.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=70 style='width: 47.7pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=27 style='width: 16.2pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl6513844 width=23 style='width: 15.3pt'></td>                                                                                                    \n");
            htmlStr.Append("				<td class=xl8713844>Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "  </td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl8513844>&nbsp;</td>                                                                                                                             \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=21 style='mso-height-source: userset; height: 18.00pt'>                                                                                             \n");
            htmlStr.Append("				<td height=21 class=xl8813844 style='height: 18.00pt'>Tra                                                                                                  \n");
            htmlStr.Append("					c&#7913;u t&#7841;i Website: <font class='font1113844'>" + dt.Rows[0]["WEBSITE_EI"] + "</font>                                                      \n");
            htmlStr.Append("				</td>                                                                                                                                                       \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=18 style='mso-height-source: userset; height: 18.00pt'>                                                                                             \n");
            htmlStr.Append("				<td colspan=26 height=18 class=xl9513844 width=709                                                                                                          \n");
            htmlStr.Append("					style='height: 18.00pt; width: 531pt'>(C&#7847;n ki&#7875;m                                                                                            \n");
            htmlStr.Append("					tra, &#273;&#7889;i chi&#7871;u khi l&#7853;p, giao, nh&#7853;n hóa                                                                                     \n");
            htmlStr.Append("					&#273;&#417;n)</td>                                                                                                                                     \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<tr height=20 style='mso-height-source: userset; height: 16.0pt'>                                                                                               \n");
            htmlStr.Append("				<td colspan=26 height=20 class=xl6513844 width=709                                                                                                          \n");
            htmlStr.Append("					style='height: 16.0pt; width: 531pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "</td>                                                                                                                        \n");
            htmlStr.Append("				<td class=xl1513844></td>                                                                                                                                   \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<![if supportMisalignedColumns]>                                                                                                                                \n");
            htmlStr.Append("			<tr height=0 style='display: none'>                                                                                                                             \n");
            htmlStr.Append("				<td width=30 style='width: 20.7pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=99 style='width: 66.6pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=16 style='width: 10.8pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=90 style='width: 60.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=70 style='width: 47.7pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=27 style='width: 16.2pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=13 style='width: 9pt'></td>                                                                                                                       \n");
            htmlStr.Append("				<td width=10 style='width: 7.2pt'></td>                                                                                                                     \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=13 style='width: 9pt'></td>                                                                                                                       \n");
            htmlStr.Append("				<td width=13 style='width: 9pt'></td>                                                                                                                       \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=13 style='width: 9pt'></td>                                                                                                                       \n");
            htmlStr.Append("				<td width=13 style='width: 9pt'></td>                                                                                                                       \n");
            htmlStr.Append("				<td width=13 style='width: 9pt'></td>                                                                                                                       \n");
            htmlStr.Append("				<td width=10 style='width: 7.2pt'></td>                                                                                                                     \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=23 style='width: 15.3pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=26 style='width: 17.1pt'></td>                                                                                                                    \n");
            htmlStr.Append("				<td width=6 style='width: 3.6pt'></td>                                                                                                                      \n");
            htmlStr.Append("			</tr>                                                                                                                                                           \n");
            htmlStr.Append("			<![endif]>                                                                                                                                                      \n");
            htmlStr.Append("		</table>                                                                                                                                                            \n");
            htmlStr.Append("	</div>                                                                                                                                                                  \n");
            htmlStr.Append("	</section>                                                                                                                                                              \n");
            htmlStr.Append("</body>                                                                                                                                                                     \n");
            htmlStr.Append("</html>                                                                                                                                                                     \n");

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
                    rtnf = rtnf + " Và" + strTmp + " Xu";
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